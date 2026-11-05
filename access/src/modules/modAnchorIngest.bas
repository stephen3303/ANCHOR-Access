Option Compare Binary
Option Explicit

' ANCHOR (Access edition) - mail ingestion, a prototype (Stephen,
' 2026-11-05): received record traffic comes into tblMessages from a
' dedicated mailbox read through Outlook (late bound) or from .txt files
' dropped in a folder. Each text goes through IngestRaw, a port of the web
' app's parseNavalMessage (src/lib/naval.ts) and ingest (src/lib/server/
' messages.ts), and ends as one of four results:
'   stored     parsed and added to Messages (with its sections and
'              addressees); a NARR directive changes the message it names
'   repeat     its serial is held with the same text (same SHA-256): logged,
'              nothing stored
'   conflict   its serial is held with different text: refused and logged,
'              for an admin to look at
'   rejected   the text doesn't parse, its type is unknown, or it is marked
'              above the highest classification this data file takes
'              (IngestMaxClassification, never above SECRET); nothing stored
' Every attempt adds one tblIngestLog row and one audit row, in the same
' transaction as anything it stores. The text stored is the raw text with its
' line breaks made CRLF and trailing spaces and blank end lines removed
' (IngNormRaw), so the same message hashes the same whether it came as a
' Windows file, a Unix file or a mail body; tblMessages.RawText keeps it
' and Sha256 is its hash.
'
' The layout read (a simplified GENADMIN, as naval.ts documents it):
'   O 241830Z OCT 26                 precedence prosign (Z O P R) and DTG
'   FM COMSUBLANT N3                 originator; Source is its first word
'   TO A, B                          addressees, by comma; a line under TO or
'   INFO C                           INFO with no keyword adds to it
'   BT
'   UNCLAS                           classification (see IngClassOf)
'   SUBJ/...//                       sets end with //
'   MSGID/GENADMIN/COMSUBLANT N3/COMSUBLANT 078/26//   type / originator / serial
'   REF/A/MSG/COMSUBLANT 061/26//
'   NARR/THIS MSG SUPERSEDES REF A.//
'   ACK/261830Z OCT 26//
'   POC/...//
'   RMKS/1. PURPOSE. TEXT ...        numbered paragraphs, A. items, to BT
'   BT
' The MSGID type must be an active type in tblMessageTypes, else the text is
' rejected "unknown type" for an admin to add the type and send it again.
' GENADMIN, the general narrative format, takes the type its serial names,
' as the web app does (NAVADMIN 236/26 is a NAVADMIN). Where naval.ts
' refuses unnumbered remarks, this takes them as one "Remarks" section,
' which is how a one-paragraph message drafted in ANCHOR reads. A line
' carrying a demonstration marking between the classification line and the
' first set (modDrafting's test-marking line) is passed over.
'
' Lifecycle directives (NARR "THIS MSG <verb> REF <letter>"): SUPERSEDES and
' CORRECTS make the message named SUPERSEDED, CANCELS makes it CANCELLED,
' AMENDS makes it AMENDED, each with ReplacedBy this message's serial; a
' CORRECTS message is itself CORRECTED, any other ACTIVE. A message named
' that ANCHOR doesn't hold is noted in the log and nothing else changes.
'
' The feeders (IngestCheckFolder, IngestCheckMailbox) move each file or mail
' item to an Ingested or Rejected subfolder they create; they never delete
' one. A rejected file gets a "<name>.reason.txt" beside it. Each run's
' counts are kept in tblSettings IngestLastRun for the Admin page.
'
' Every run and every attempt checks, at that moment in the data file, that
' the signed-in person is a System Admin (RequireAnchorAdmin).
'
' Owner: mail ingestion. Depends on: modAudit, modErrors, modGlobals,
' modSignIn, modUI (ClipText); the kit's modCrypto (Sha256Hex), modKitBackend
' (KitBackendPath), modKitLocks (KitExecute), modKitUtil (EnsureFolder,
' FolderExists, FileExists).

' The four results IngestRaw returns (tblIngestLog.Result).
Public Const INGEST_STORED As String = "stored"
Public Const INGEST_REPEAT As String = "repeat"
Public Const INGEST_CONFLICT As String = "conflict"
Public Const INGEST_REJECTED As String = "rejected"

' The settings (tblSettings SettingKey).
Public Const INGEST_SET_FOLDER As String = "IngestDropFolder"
Public Const INGEST_SET_MAILBOX As String = "IngestMailbox"
Public Const INGEST_SET_MAILFOLDER As String = "IngestFolder"
Public Const INGEST_SET_CLASS As String = "IngestDefaultClass"
Public Const INGEST_SET_MAX As String = "IngestMaxClassification"
Public Const INGEST_SET_MINUTES As String = "IngestTimerMinutes"
Public Const INGEST_SET_LASTRUN As String = "IngestLastRun"

Private Const MAX_CHARS As Long = 65536             ' the web app's 64 KB limit
Private Const READ_LIMIT As Long = 1048576          ' bytes a feeder reads from one file at most
Private Const PR_INTERNET_MESSAGE_ID As String = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
Private Const OL_MAIL_ITEM As Long = 43             ' olMail
Private Const OL_FOLDER_INBOX As Long = 6           ' olFolderInbox
Private Const MONTHS As String = "JANFEBMARAPRMAYJUNJULAUGSEPOCTNOVDEC"

Private m_lastReason As String      ' why the last IngestRaw ended as it did ("" for a plain store)
Private m_lastMsgId As Long         ' the message the last IngestRaw stored or found (0: none)
Private m_runFeeder As String       ' the run's counts (IngStartRun, IngFinishRun)
Private m_runStored As Long
Private m_runRepeat As Long
Private m_runConflict As Long
Private m_runRejected As Long
Private m_runErrors As Long
Private m_runError As String        ' the first error a run met, for its summary

' ---------------------------------------------------------------------------
' One message
' ---------------------------------------------------------------------------

' Parse, check and store one raw message, logging the attempt.
'   raw          the message text as it arrived
'   source       the feeder: "Folder", "Outlook" or another short word
'   msgIdHeader  the mail's Internet Message-ID, "" for none
'   fileOrItem   the file name or mail item it came from, for the log
' Returns INGEST_STORED, INGEST_REPEAT, INGEST_CONFLICT or INGEST_REJECTED;
'   IngestLastReason says why, IngestLastMessageID which message it stored
'   or matched.
' Errors: ERR_NOT_ALLOWED unless the signed-in person is a System Admin now;
'   5 when the data file has no tblIngestLog (made before ingestion); the
'   engine's (3022 when another copy stored the same serial at the same
'   moment: nothing was written, and the next attempt finds it).
' Side effects: one transaction with the message rows (stored), the change
'   to the message a directive names, one tblIngestLog row and one audit row.
Public Function IngestRaw(ByVal raw As String, ByVal source As String, Optional ByVal msgIdHeader As String = "", _
        Optional ByVal fileOrItem As String = "") As String
    Dim ws As DAO.Workspace, inTrans As Boolean, p As Object, why As String, txt As String, hash As String
    Dim result As String, ref As String, msgId As Long, classif As String, auditResult As String, resource As String
    m_lastReason = "": m_lastMsgId = 0
    RequireAnchorAdmin
    If Not IngestReady() Then Err.Raise 5, "modAnchorIngest.IngestRaw", _
        "This data file has no ingest log. Make the data file again (backend\new-backend.ps1) to add mail ingestion."
    txt = IngNormRaw(raw)
    hash = LCase$(Sha256Hex(txt))
    If Len(txt) = 0 Then
        why = "The text is empty."
    ElseIf Len(txt) > MAX_CHARS Then
        why = "The text is longer than 64 KB."
    Else
        Set p = IngParse(txt, why, ref)
    End If
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    If p Is Nothing Then
        result = INGEST_REJECTED
    Else
        result = IngDuplicate(ref, p("msg_key"), hash, why, msgId)
        If Len(result) = 0 Then
            msgId = IngStore(p, txt, hash, why)
            result = INGEST_STORED
            classif = p("classification")
        End If
    End If
    KitExecute "INSERT INTO tblIngestLog (LoggedAt, Workstation, UserName, Feeder, EmailMessageId, Reference, Sha256, Result, Reason, " & _
        "FileOrItem, MessageID) VALUES (Now(), " & IngSqlText(Left$(Environ$("COMPUTERNAME"), 64)) & ", " & _
        IngSqlText(Left$(SignedInName(), 100)) & ", " & Q(Left$(source, 20)) & ", " & IngSqlText(Left$(msgIdHeader, 255)) & ", " & _
        IngSqlText(Left$(ref, 60)) & ", " & Q(hash) & ", " & Q(result) & ", " & IngSqlText(Left$(why, 255)) & ", " & _
        IngSqlText(Left$(fileOrItem, 255)) & ", " & IIf(msgId = 0, "Null", msgId) & ")"
    ' A refused text's own marking is never written to the audit trail.
    If Len(classif) = 0 Then classif = "UNCLASSIFIED"
    Select Case result
        Case INGEST_STORED: auditResult = "Success"
        Case INGEST_REPEAT: auditResult = "Warning"
        Case Else: auditResult = "Failed"
    End Select
    If Len(ref) > 0 Then resource = "Message (" & ref & ")" Else resource = "Ingest (" & ClipText(fileOrItem, 120) & ")"
    AuditLog "Ingest", resource, StrConv(result, vbProperCase) & " from the " & source & " feeder" & _
        IIf(Len(fileOrItem) > 0, " (" & fileOrItem & ")", "") & "." & IIf(Len(why) > 0, " " & why, ""), classif, auditResult
    ws.CommitTrans: inTrans = False
    m_lastReason = why: m_lastMsgId = msgId
    IngestRaw = result
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modAnchorIngest.IngestRaw", fileOrItem
End Function

' Why the last IngestRaw ended as it did: the reason a text was refused, or
' a note on what a stored one changed. "" when there is nothing to say.
Public Function IngestLastReason() As String
    IngestLastReason = m_lastReason
End Function

' The message the last IngestRaw stored, or the one a repeat or conflict
' matched; 0 for none.
Public Function IngestLastMessageID() As Long
    IngestLastMessageID = m_lastMsgId
End Function

' True when a data file with the ingest log is linked.
Public Function IngestReady() As Boolean
    If Not DataLinked() Then Exit Function
    IngestReady = TableExists("tblIngestLog")
End Function

' The serial is held already: INGEST_REPEAT (same text) or INGEST_CONFLICT
' (other text), with why and the held message's id; "" when it is new. The
' held text is hashed again as IngNormRaw makes it, so a message seeded or
' stored before keeps matching whatever its line breaks.
Private Function IngDuplicate(ByVal ref As String, ByVal msgKey As String, ByVal hash As String, ByRef why As String, _
        ByRef msgId As Long) As String
    Dim rs As DAO.Recordset, same As Boolean
    Set rs = CurrentDb.OpenRecordset("SELECT MessageID, Reference, Sha256, RawText FROM tblMessages WHERE Reference=" & Q(ref) & _
        " OR MsgKey=" & Q(msgKey), dbOpenSnapshot)
    If Not rs.EOF Then
        msgId = rs!MessageID
        same = (LCase$(Nz(rs!Sha256, "")) = hash)
        If Not same Then same = (LCase$(Sha256Hex(IngNormRaw(Nz(rs!RawText, "")))) = hash)
        If same Then
            IngDuplicate = INGEST_REPEAT
            why = Nz(rs!Reference, ref) & " is already held with the same text. Nothing was stored."
        Else
            IngDuplicate = INGEST_CONFLICT
            why = Nz(rs!Reference, ref) & " is already held with different text. Refused: each serial is stored once; " & _
                "a correction or amendment comes as new traffic."
        End If
    End If
    rs.Close
End Function

' Add parsed message p to tblMessages with its sections and addressees, and
' apply its directive. Returns the new MessageID; outNote says what the
' directive did.
Private Function IngStore(ByVal p As Object, ByVal txt As String, ByVal hash As String, ByRef outNote As String) As Long
    Dim rs As DAO.Recordset, msgRow As Long, v As Variant, i As Long, ref As String, tl As String, n As Long
    ref = p("reference")
    Set rs = CurrentDb.OpenRecordset("tblMessages", dbOpenDynaset, dbAppendOnly)
    rs.AddNew
    rs!MsgKey = p("msg_key")
    rs!Dtg = p("dtg")
    rs!ReceivedOn = p("received")
    rs!Subject = Left$(p("subject"), 255)
    rs!Reference = ref
    rs!MsgType = p("msg_type")
    rs!Originator = Left$(p("originator"), 60)
    rs!Source = Left$(p("source"), 40)
    rs!ToLine = Left$(p("to_line"), 255)
    rs!InfoLine = Left$(p("info_line"), 255)
    rs!Precedence = p("precedence")
    rs!Classification = p("classification")
    rs!Importance = p("importance")
    rs!Lifecycle = p("lifecycle")
    If Len(p("target")) > 0 Then rs!RelatesTo = Left$(p("target"), 60)
    rs!AckRequired = Not IsNull(p("ack_due"))
    rs!AckDue = p("ack_due")
    rs!C2oixId = IngNextC2oixId(p("received"))
    rs!Sha256 = hash
    rs!SizeKb = (Len(txt) + 1023) \ 1024
    rs!MsgClass = IngestDefaultClass()
    rs!Retention = "Permanent"
    If Len(p("related")) > 0 Then rs!Related = Left$(p("related"), 255)
    rs!RawText = txt
    rs!BodyText = p("body")
    rs.Update
    rs.Bookmark = rs.LastModified
    msgRow = rs!MessageID
    rs.Close
    Set rs = CurrentDb.OpenRecordset("tblMessageSections", dbOpenDynaset, dbAppendOnly)
    For Each v In p("sections")
        i = i + 1
        rs.AddNew
        rs!MessageID = msgRow
        rs!SortOrder = i
        rs!Heading = Left$(v(0), 100)
        rs!Body = v(1)
        rs.Update
    Next
    rs.Close
    For Each v In p("addressees")
        IngAddAddressee msgRow, CStr(v(0)), CStr(v(1))
    Next
    If Len(p("directive")) > 0 Then
        tl = IngTargetLifecycle(p("directive"))
        n = KitExecute("UPDATE tblMessages SET Lifecycle=" & Q(tl) & ", ReplacedBy=" & Q(ref) & " WHERE Reference=" & Q(p("target")) & _
            " AND MessageID<>" & msgRow)
        If n = 0 Then
            outNote = "It " & LCase$(p("directive")) & " " & p("target") & ", which ANCHOR doesn't hold; no other message changed."
        Else
            outNote = p("target") & " is now " & tl & " (" & LCase$(p("directive")) & " by " & ref & ")."
        End If
    End If
    IngStore = msgRow
End Function

' An addressee, a group when it names an address group (its name as the
' group spells it), else a command.
Private Sub IngAddAddressee(ByVal msgRow As Long, ByVal nm As String, ByVal role As String)
    Dim grp As Variant
    grp = DLookup("GroupName", "tblAddressGroups", "GroupName=" & Q(nm))
    KitExecute "INSERT INTO tblAddressees (MessageID, AddrName, AddrRole, AddrKind) VALUES (" & msgRow & ", " & _
        Q(Left$(Nz(grp, nm), 100)) & ", " & Q(role) & ", " & IIf(IsNull(grp), "'command'", "'group'") & ")"
End Sub

' What a directive does to the message it names.
Private Function IngTargetLifecycle(ByVal kind As String) As String
    Select Case kind
        Case "CANCELS": IngTargetLifecycle = "CANCELLED"
        Case "AMENDS": IngTargetLifecycle = "AMENDED"
        Case Else: IngTargetLifecycle = "SUPERSEDED"       ' SUPERSEDES, CORRECTS
    End Select
End Function

' The next C2OIX id, "C2OIX-26-048731": the year of the DTG and one more
' than the highest number held (the web app numbers them in one sequence).
Private Function IngNextC2oixId(ByVal received As Date) As String
    Dim rs As DAO.Recordset, s As String, mx As Long, k As Long
    Set rs = CurrentDb.OpenRecordset("SELECT C2oixId FROM tblMessages WHERE C2oixId Like 'C2OIX-*'", dbOpenSnapshot)
    Do While Not rs.EOF
        s = Nz(rs!C2oixId, "")
        k = InStrRev(s, "-")
        If k > 0 Then
            If IsNumeric(Mid$(s, k + 1)) Then
                If CLng(Val(Mid$(s, k + 1))) > mx Then mx = CLng(Val(Mid$(s, k + 1)))
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    IngNextC2oixId = "C2OIX-" & Format$(received, "yy") & "-" & Format$(mx + 1, "000000")
End Function

' ---------------------------------------------------------------------------
' The parser (naval.ts parseNavalMessage)
' ---------------------------------------------------------------------------

' Parse normalised text txt (IngNormRaw). Returns a Scripting.Dictionary of
' the message's fields, or Nothing with why set to a reason an operator can
' act on. ref is set as soon as the serial is read, for the log.
Private Function IngParse(ByVal txt As String, ByRef why As String, ByRef ref As String) As Object
    Dim lines() As String, ub As Long, i As Long, t As String, u As String, v As Variant, k As Long
    Dim prec As String, dtg As String, received As Variant, fm As String, toL As String, infoL As String, lastHdr As String
    Dim cls As String, sets As Collection, rmks As Collection, cur As String, inRmks As Boolean
    Dim body As String, subj As String, orig As String, serial As String, fmt As String, narr As String, poc As String
    Dim ackDue As Variant, refs As Collection, rest As String, letter As String, kind As Variant, phrase As String
    Dim dirKind As String, target As String, msgType As String, p As Object, related As String
    Dim heads() As String, bodies() As String, nSec As Long, head As String, para As String, sections As Collection, adds As Collection
    Dim lifecycle As String, rmksText As String

    lines = Split(Replace(txt, vbCrLf, vbLf), vbLf)
    ub = UBound(lines)
    ackDue = Null

    ' Line 1: precedence prosign and DTG.
    u = IngSqueeze(UCase$(Trim$(lines(0))))
    If Len(u) < 3 Or InStr(1, "ZOPR", Left$(u, 1)) = 0 Or Mid$(u, 2, 1) <> " " Then
        why = "Line 1 must be a precedence prosign and date-time group, e.g. ""O 241830Z OCT 26""."
        Exit Function
    End If
    dtg = Mid$(u, 3)
    If Not dtg Like "######Z [A-Z][A-Z][A-Z] ##" Then
        why = "Line 1 must be a precedence prosign and date-time group, e.g. ""O 241830Z OCT 26""."
        Exit Function
    End If
    Select Case Left$(u, 1)
        Case "Z": prec = "FLASH"
        Case "O": prec = "IMMEDIATE"
        Case "P": prec = "PRIORITY"
        Case Else: prec = "ROUTINE"
    End Select
    received = IngDtgDate(dtg)
    If IsNull(received) Then why = """" & dtg & """ is not a valid date-time group.": Exit Function

    ' Heading lines up to BT.
    For i = 1 To ub
        t = Trim$(lines(i))
        u = UCase$(t)
        If u = "BT" Then Exit For
        If Left$(u, 3) = "FM " Then
            fm = Trim$(Mid$(t, 4)): lastHdr = "FM"
        ElseIf Left$(u, 3) = "TO " Then
            toL = Trim$(Mid$(t, 4)): lastHdr = "TO"
        ElseIf Left$(u, 5) = "INFO " Then
            infoL = Trim$(Mid$(t, 6)): lastHdr = "INFO"
        ElseIf Len(t) > 0 And lastHdr = "TO" Then
            toL = toL & IIf(Len(toL) > 0, ", ", "") & t       ' one address a line, as a drafted message has them
        ElseIf Len(t) > 0 And lastHdr = "INFO" Then
            infoL = infoL & IIf(Len(infoL) > 0, ", ", "") & t
        End If
    Next
    If i > ub Then why = "Missing the BT line that opens the text.": Exit Function
    If Len(fm) = 0 Then why = "Missing the FM line.": Exit Function
    If Len(toL) = 0 Then why = "Missing the TO line.": Exit Function
    i = i + 1

    ' The classification line, then the highest level this data file takes.
    If i <= ub Then cls = IngClassOf(UCase$(Trim$(lines(i))), why) Else cls = IngClassOf("", why)
    If Len(cls) = 0 Then Exit Function
    If IngClassRank(cls) > IngClassRank(IngestMaxClassification()) Then
        why = "The text is marked " & cls & ", above the highest level this data file takes (" & IngestMaxClassification() & _
            "). Refused and not stored."
        Exit Function
    End If
    i = i + 1

    ' Sets run until a line ending in //; RMKS is the last and runs to BT.
    Set sets = New Collection
    Set rmks = New Collection
    Do While i <= ub
        t = Trim$(lines(i))
        If UCase$(t) = "BT" Then Exit Do
        If inRmks Then
            rmks.Add lines(i)
        ElseIf Len(cur) = 0 And UCase$(Left$(t, 5)) = "RMKS/" Then
            inRmks = True
            rmks.Add Mid$(t, 6)
        ElseIf Len(t) = 0 Then
            ' a blank line between sets
        ElseIf Len(cur) = 0 And sets.Count = 0 And Not IngIsSetStart(t) And InStr(1, UCase$(t), "MARKING") > 0 Then
            ' the demonstration marking line under the classification
        Else
            If Len(cur) > 0 And IngIsSetStart(t) Then
                why = "The set starting """ & Left$(cur, 30) & """ is not closed with //."
                Exit Function
            End If
            cur = cur & IIf(Len(cur) > 0, " ", "") & t
            If Right$(t, 2) = "//" Then sets.Add cur: cur = ""
        End If
        i = i + 1
    Loop
    If Len(cur) > 0 Then why = "The set starting """ & Left$(cur, 30) & """ is not closed with //.": Exit Function
    If i > ub Then why = "Missing the BT line that closes the text.": Exit Function

    orig = fm
    Set refs = New Collection
    For Each v In sets
        body = Left$(v, Len(v) - 2)
        u = UCase$(body)
        If Left$(u, 5) = "SUBJ/" Then
            subj = Trim$(Mid$(body, 6))
        ElseIf Left$(u, 6) = "MSGID/" Then
            ' MSGID/<type>/<originator>/<serial>: the serial may hold a slash
            ' ("COMSUBLANT 078/26"), so it takes the rest of the set.
            rest = Mid$(body, 7)
            k = InStr(rest, "/")
            If k > 0 Then
                fmt = UCase$(Trim$(Left$(rest, k - 1)))
                rest = Mid$(rest, k + 1)
                k = InStr(rest, "/")
                If k > 0 Then
                    If Len(Trim$(Left$(rest, k - 1))) > 0 Then orig = Trim$(Left$(rest, k - 1))
                    serial = Trim$(Mid$(rest, k + 1))
                End If
            End If
        ElseIf Left$(u, 4) = "REF/" Then
            rest = Mid$(body, 5)
            If Left$(rest, 1) Like "[A-Z]" And Mid$(rest, 2, 1) = "/" Then
                letter = Left$(rest, 1)
                rest = Mid$(rest, 3)
                k = InStr(rest, "/")
                If k > 0 Then refs.Add Array(letter, IngNormRef(Mid$(rest, k + 1)))
            End If
        ElseIf Left$(u, 5) = "NARR/" Then
            narr = Mid$(body, 6)
        ElseIf Left$(u, 4) = "ACK/" Then
            ackDue = IngDtgDate(Mid$(body, 5))
            If IsNull(ackDue) Then why = "ACK due """ & Trim$(Mid$(body, 5)) & """ is not a valid date-time group.": Exit Function
        ElseIf Left$(u, 4) = "POC/" Then
            poc = Trim$(Mid$(body, 5))
        End If
    Next
    If Len(serial) > 0 Then ref = IngNormRef(serial)
    If Len(subj) = 0 Then why = "Missing the SUBJ/ set.": Exit Function
    If Len(serial) = 0 Then
        why = "Missing the message serial in MSGID, e.g. ""MSGID/GENADMIN/COMSUBLANT N3/COMSUBLANT 078/26//""."
        Exit Function
    End If
    If Not IngSerialOk(ref) Then why = """" & serial & """ is not a message serial like ""NAVADMIN 236/26"".": Exit Function
    If Len(ref) > 60 Then why = "The serial is longer than 60 characters.": Exit Function
    If Not inRmks Then why = "Missing the RMKS/ remarks.": Exit Function

    ' The type: the MSGID's, which must be a type this data file has.
    If Len(fmt) = 0 Then
        why = "unknown type: the MSGID/ set names no message type."
        Exit Function
    End If
    msgType = IngActiveType(fmt)
    If Len(msgType) = 0 Then
        why = "unknown type: " & fmt & " is not an active message type here. A System Admin adds it under Admin: " & _
            "Message types, then the text is sent again."
        Exit Function
    End If
    If msgType = "GENADMIN" Then
        ' The general narrative format: the serial names the type, as the
        ' web app's typeOf reads it.
        u = IngActiveType(IngTypeOfSerial(ref, orig))
        If Len(u) > 0 Then msgType = u
    End If

    ' NARR "THIS MSG <verb> REF <letter>" names a REF set.
    u = UCase$(narr)
    For Each kind In Array("SUPERSEDES", "CANCELS", "CORRECTS", "AMENDS")
        phrase = "THIS MSG " & kind & " REF "
        k = InStr(1, u, phrase)
        If k > 0 Then
            letter = Mid$(u, k + Len(phrase), 1)
            If letter Like "[A-Z]" And Not Mid$(u, k + Len(phrase) + 1, 1) Like "[A-Z0-9_]" Then
                dirKind = kind
                For Each v In refs
                    If v(0) = letter Then target = v(1): Exit For
                Next
                If Len(target) = 0 Then why = "NARR names REF " & letter & " but there is no REF/" & letter & "/ set.": Exit Function
                Exit For
            End If
        End If
    Next

    ' Remarks: "N. HEADING. TEXT" paragraphs, continuation lines and "A. ITEM"
    ' sub-items. Text before the first numbered paragraph is "Remarks".
    For Each v In rmks
        t = Trim$(IngNoTerminator(Trim$(v)))
        rmksText = rmksText & IIf(Len(rmksText) > 0, vbCrLf, "") & IngNoTerminator(RTrim$(v))
        If Len(t) > 0 Then
            If IngParaParts(t, head, para) Then
                nSec = nSec + 1
                ReDim Preserve heads(1 To nSec): ReDim Preserve bodies(1 To nSec)
                heads(nSec) = IngTitleCase(head): bodies(nSec) = para
            ElseIf nSec = 0 Then
                nSec = 1
                ReDim heads(1 To 1): ReDim bodies(1 To 1)
                heads(1) = "Remarks": bodies(1) = t
            ElseIf IngItemText(t, para) Then
                bodies(nSec) = bodies(nSec) & IIf(Len(bodies(nSec)) > 0, vbCrLf, "") & para
            ElseIf Len(bodies(nSec)) > 0 Then
                bodies(nSec) = bodies(nSec) & " " & t
            Else
                bodies(nSec) = t
            End If
        End If
    Next
    If nSec = 0 Then why = "The remarks are empty.": Exit Function
    Set sections = New Collection
    For k = 1 To nSec
        sections.Add Array(heads(k), bodies(k))
    Next
    If Len(poc) > 0 Then sections.Add Array("Point of Contact", poc)

    Set adds = New Collection
    For Each v In Split(toL, ",")
        If Len(Trim$(v)) > 0 Then adds.Add Array(Trim$(v), "TO")
    Next
    For Each v In Split(infoL, ",")
        If Len(Trim$(v)) > 0 Then adds.Add Array(Trim$(v), "INFO")
    Next
    For Each v In refs
        related = related & IIf(Len(related) > 0, "; ", "") & v(1)
    Next
    If dirKind = "CORRECTS" Then lifecycle = "CORRECTED" Else lifecycle = "ACTIVE"

    Set p = CreateObject("Scripting.Dictionary")
    p("precedence") = prec
    p("dtg") = dtg
    p("received") = received
    p("originator") = orig
    p("source") = UCase$(Split(Trim$(orig) & " ", " ")(0))
    p("to_line") = toL
    p("info_line") = infoL
    p("classification") = cls
    p("reference") = ref
    p("msg_key") = IngKeyOf(ref)
    p("msg_type") = msgType
    p("subject") = subj
    p("related") = related
    p("directive") = dirKind
    p("target") = target
    p("ack_due") = ackDue
    p("lifecycle") = lifecycle
    Select Case prec
        Case "FLASH", "IMMEDIATE": p("importance") = "High"
        Case "PRIORITY": p("importance") = "Medium"
        Case Else: p("importance") = "Low"
    End Select
    p("body") = rmksText
    Set p("sections") = sections
    Set p("addressees") = adds
    Set IngParse = p
End Function

' A line without the // that ends a set.
Private Function IngNoTerminator(ByVal s As String) As String
    If Right$(s, 2) = "//" Then IngNoTerminator = RTrim$(Left$(s, Len(s) - 2)) Else IngNoTerminator = s
End Function

' The classification a text classification line names (spaced letters, as
' NTP 3 spells a classified one, read as one word), or "" with why. Above
' SECRET, or a line not recognised, is refused.
Private Function IngClassOf(ByVal classLine As String, ByRef why As String) As String
    Dim k As String
    k = Replace(Replace(classLine, " ", ""), vbTab, "")
    If InStr(1, k, "TOPSECRET") > 0 Or k = "TS" Or Left$(k, 4) = "TS//" Or InStr(1, k, "//SCI") > 0 Or InStr(1, k, "//SI") > 0 _
            Or InStr(1, k, "//TK") > 0 Or InStr(1, k, "//HCS") > 0 Then
        why = "The text is marked above SECRET. ANCHOR refuses it and stores nothing."
        Exit Function
    End If
    Select Case k
        Case "UNCLAS", "UNCLASSIFIED": IngClassOf = "UNCLASSIFIED"
        Case "UNCLAS//CUI", "UNCLASSIFIED//CUI", "CUI": IngClassOf = "CUI"
        Case "CONFIDENTIAL": IngClassOf = "CONFIDENTIAL"
        Case "SECRET": IngClassOf = "SECRET"
        Case Else: why = "Unknown classification line """ & ClipText(classLine, 40) & """."
    End Select
End Function

' 0 UNCLASSIFIED, 1 CUI, 2 CONFIDENTIAL, 3 SECRET (anything else counts as SECRET).
Private Function IngClassRank(ByVal cls As String) As Long
    Select Case UCase$(cls)
        Case "UNCLASSIFIED": IngClassRank = 0
        Case "CUI": IngClassRank = 1
        Case "CONFIDENTIAL": IngClassRank = 2
        Case Else: IngClassRank = 3
    End Select
End Function

' A line that opens a set: seen before the last set's // means that set was left open.
Private Function IngIsSetStart(ByVal t As String) As Boolean
    Dim w As Variant, u As String
    u = UCase$(t)
    For Each w In Array("SUBJ/", "MSGID/", "REF/", "NARR/", "ACK/", "POC/", "RMKS/", "AMPN/", "GENTEXT/")
        If Left$(u, Len(w)) = w Then IngIsSetStart = True: Exit Function
    Next
End Function

' "241830Z OCT 26" as a date-time (UTC as written), or Null when it isn't a
' valid DTG.
Private Function IngDtgDate(ByVal s As String) As Variant
    Dim d As Long, hh As Long, mm As Long, mon As Long, yy As Long, dt As Date
    IngDtgDate = Null
    s = IngSqueeze(UCase$(Trim$(s)))
    If Not s Like "######Z [A-Z][A-Z][A-Z] ##" Then Exit Function
    d = CLng(Left$(s, 2)): hh = CLng(Mid$(s, 3, 2)): mm = CLng(Mid$(s, 5, 2)): yy = CLng(Right$(s, 2))
    mon = InStr(1, MONTHS, Mid$(s, 9, 3))
    If mon = 0 Or (mon - 1) Mod 3 <> 0 Then Exit Function
    mon = (mon + 2) \ 3
    If d < 1 Or d > 31 Or hh > 23 Or mm > 59 Then Exit Function
    dt = DateSerial(2000 + yy, mon, d)
    If Day(dt) <> d Or Month(dt) <> mon Then Exit Function
    IngDtgDate = dt + TimeSerial(hh, mm, 0)
End Function

' "N. HEADING. TEXT" as heading and text (naval.ts: /^(\d+)\.\s+([^.]+)\.(?:\s+(.*))?$/).
Private Function IngParaParts(ByVal l As String, ByRef head As String, ByRef para As String) As Boolean
    Dim k As Long, d As Long, rest As String
    k = 1
    Do While k <= Len(l)
        If Not Mid$(l, k, 1) Like "#" Then Exit Do
        k = k + 1
    Loop
    If k = 1 Or Mid$(l, k, 1) <> "." Then Exit Function
    k = k + 1
    If Mid$(l, k, 1) <> " " And Mid$(l, k, 1) <> vbTab Then Exit Function
    Do While Mid$(l, k, 1) = " " Or Mid$(l, k, 1) = vbTab
        k = k + 1
    Loop
    d = InStr(k, l, ".")
    If d <= k Then Exit Function
    rest = Mid$(l, d + 1)
    If Len(rest) > 0 And Left$(rest, 1) <> " " And Left$(rest, 1) <> vbTab Then Exit Function
    head = Trim$(Mid$(l, k, d - k))
    para = Trim$(rest)
    IngParaParts = True
End Function

' "A. ITEM" as its text (naval.ts: /^([A-Z])\.\s+(.*)$/).
Private Function IngItemText(ByVal l As String, ByRef para As String) As Boolean
    If Not Left$(l, 1) Like "[A-Z]" Or Mid$(l, 2, 1) <> "." Then Exit Function
    If Mid$(l, 3, 1) <> " " And Mid$(l, 3, 1) <> vbTab Then Exit Function
    para = Trim$(Mid$(l, 3))
    IngItemText = True
End Function

' "POINT OF CONTACT" -> "Point Of Contact": each letter after a non-word
' character in capitals, the rest small (naval.ts titleCase).
Private Function IngTitleCase(ByVal s As String) As String
    Dim i As Long, ch As String, prevWord As Boolean, out As String
    s = LCase$(s)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[a-z]" And Not prevWord Then ch = UCase$(ch)
        prevWord = (ch Like "[A-Za-z0-9_]")
        out = out & ch
    Next
    IngTitleCase = out
End Function

' A serial like "NAVADMIN 236/26", "COMSUBFOR INST 3500.2" or "USFF 2026-11"
' (naval.ts: /^[A-Z][A-Z .]* [\d.-]+(\/\d{2})?$/i), already upper case.
Private Function IngSerialOk(ByVal s As String) As Boolean
    Dim k As Long, prefix As String, num As String, i As Long, ch As String
    k = InStrRev(s, " ")
    If k < 2 Then Exit Function
    prefix = Left$(s, k - 1): num = Mid$(s, k + 1)
    If Not Left$(prefix, 1) Like "[A-Z]" Then Exit Function
    For i = 2 To Len(prefix)
        If Not Mid$(prefix, i, 1) Like "[A-Z .]" Then Exit Function
    Next
    If num Like "*/##" Then num = Left$(num, Len(num) - 3)
    If Len(num) = 0 Then Exit Function
    For i = 1 To Len(num)
        ch = Mid$(num, i, 1)
        If Not ch Like "[0-9.-]" Then Exit Function
    Next
    IngSerialOk = True
End Function

' The type a serial names (naval.ts typeOf): SAFETY MSG when the originator
' is a safety office, else by the serial's first word(s), FLEET MSG for any
' other.
Private Function IngTypeOfSerial(ByVal ref As String, ByVal orig As String) As String
    Dim v As Variant, p() As String
    If IngHasWord(UCase$(orig), "SAFETY") Then IngTypeOfSerial = "SAFETY MSG": Exit Function
    For Each v In Array("NAVADMIN|NAVADMIN", "ALNAV|ALNAV", "OPNAVINST|OPNAVINST", "MARADMIN|MARADMIN", "NAVSEAINST|NAVSEAINST", _
            "COMSUBFOR INST|COMSUBFOR INST", "COMSUBLANT|COMSUBLANT MSG", "SAFETY|SAFETY MSG")
        p = Split(v, "|")
        If Left$(ref, Len(p(0))) = p(0) And Not Mid$(ref, Len(p(0)) + 1, 1) Like "[A-Z0-9_]" Then
            IngTypeOfSerial = p(1)
            Exit Function
        End If
    Next
    IngTypeOfSerial = "FLEET MSG"
End Function

' True when s (upper case) holds w as a whole word.
Private Function IngHasWord(ByVal s As String, ByVal w As String) As Boolean
    Dim k As Long
    k = InStr(1, s, w)
    Do While k > 0
        If (k = 1 Or Not Mid$(s, k - 1, 1) Like "[A-Z0-9_]") And Not Mid$(s, k + Len(w), 1) Like "[A-Z0-9_]" Then
            IngHasWord = True
            Exit Function
        End If
        k = InStr(k + 1, s, w)
    Loop
End Function

' The type code as tblMessageTypes spells it when it is an active type, else "".
Private Function IngActiveType(ByVal code As String) As String
    If Len(code) = 0 Then Exit Function
    IngActiveType = Nz(DLookup("TypeCode", "tblMessageTypes", "TypeCode=" & Q(code) & " AND IsActive"), "")
End Function

' Serials compare without regard to case or spacing ("Alnav  076/26" is ALNAV 076/26).
Private Function IngNormRef(ByVal s As String) As String
    IngNormRef = UCase$(IngSqueeze(Trim$(s)))
End Function

' The message key of a serial: "NAVADMIN 236/26" -> "navadmin-236-26" (naval.ts idFor).
Private Function IngKeyOf(ByVal ref As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(ref)
        ch = LCase$(Mid$(ref, i, 1))
        If ch Like "[a-z0-9]" Then
            out = out & ch
        ElseIf Right$(out, 1) <> "-" And Len(out) > 0 Then
            out = out & "-"
        End If
    Next
    If Right$(out, 1) = "-" Then out = Left$(out, Len(out) - 1)
    IngKeyOf = Left$(out, 40)
End Function

' Runs of spaces and tabs as one space.
Private Function IngSqueeze(ByVal s As String) As String
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    IngSqueeze = s
End Function

' The text as ANCHOR stores and hashes it: no byte-order mark, CRLF line
' breaks, no spaces at line ends, no blank lines at either end.
Private Function IngNormRaw(ByVal raw As String) As String
    Dim lines() As String, i As Long, first As Long, last As Long, out As String
    If Left$(raw, 1) = ChrW(&HFEFF) Then raw = Mid$(raw, 2)
    raw = Replace(Replace(raw, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(raw, vbLf)
    first = -1
    For i = 0 To UBound(lines)
        Do While Len(lines(i)) > 0 And (Right$(lines(i), 1) = " " Or Right$(lines(i), 1) = vbTab)
            lines(i) = Left$(lines(i), Len(lines(i)) - 1)
        Loop
        If Len(lines(i)) > 0 Then
            If first < 0 Then first = i
            last = i
        End If
    Next
    If first < 0 Then Exit Function
    For i = first To last
        out = out & IIf(i > first, vbCrLf, "") & lines(i)
    Next
    IngNormRaw = out
End Function

' A text value for SQL, or Null for "".
Private Function IngSqlText(ByVal s As String) As String
    If Len(s) = 0 Then IngSqlText = "Null" Else IngSqlText = Q(s)
End Function

' ---------------------------------------------------------------------------
' Settings
' ---------------------------------------------------------------------------

' The drop folder: the IngestDropFolder setting, or the "ingest" folder
' beside the data file when it is blank. "" with no data file linked.
Public Function IngestDropFolder() As String
    Dim s As String
    s = Trim$(SettingText(INGEST_SET_FOLDER, ""))
    If Len(s) = 0 Then s = IngestDefaultDropFolder()
    If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
    IngestDropFolder = s
End Function

' The "ingest" folder beside the data file, "" with no data file linked.
Public Function IngestDefaultDropFolder() As String
    Dim be As String
    be = KitBackendPath("tblSettings")
    If Len(be) = 0 Or InStrRev(be, "\") = 0 Then Exit Function
    IngestDefaultDropFolder = Left$(be, InStrRev(be, "\")) & "ingest"
End Function

' The communications class a received message takes (IngestDefaultClass,
' A when it isn't a class this data file has).
Public Function IngestDefaultClass() As String
    Dim s As String
    s = UCase$(Trim$(SettingText(INGEST_SET_CLASS, "A")))
    If DataLinked() Then
        If DCount("*", "tblMessageClasses", "ClassCode=" & Q(s)) = 0 Then s = "A"
    End If
    IngestDefaultClass = s
End Function

' The highest classification taken (IngestMaxClassification): UNCLASSIFIED,
' CUI, CONFIDENTIAL or SECRET; SECRET when the setting is anything else.
Public Function IngestMaxClassification() As String
    Dim s As String
    s = UCase$(Trim$(SettingText(INGEST_SET_MAX, "SECRET")))
    Select Case s
        Case "UNCLASSIFIED", "CUI", "CONFIDENTIAL", "SECRET"
        Case Else: s = "SECRET"
    End Select
    IngestMaxClassification = s
End Function

' Minutes between automatic checks while the Admin page is open; 0 is off.
Public Function IngestTimerMinutes() As Long
    Dim s As String, n As Long
    s = Trim$(SettingText(INGEST_SET_MINUTES, "0"))
    If Len(s) > 0 And Len(s) <= 4 And s Like String$(Len(s), "#") Then n = CLng(s)
    If n > 1440 Then n = 0
    IngestTimerMinutes = n
End Function

' The last run's summary ("05 Nov 2026 14:02, Folder: 3 stored, ..."), or "".
Public Function IngestLastRunText() As String
    IngestLastRunText = SettingText(INGEST_SET_LASTRUN, "")
End Function

' Save the ingest settings. Returns "" when saved (or nothing changed), else
' what to fix; nothing is saved then.
'   dropFolder   "" for the folder beside the data file
'   mailbox      the mailbox as Outlook names it, or its address; "" for none
'   mailFolder   the folder in it, "Inbox" or "Inbox\ANCHOR"; "" for Inbox
'   defClass     A to E; maxClass UNCLASSIFIED, CUI, CONFIDENTIAL or SECRET
'   minutes      0 to 1440, as text
' Errors: ERR_NOT_ALLOWED unless a System Admin now; the engine's.
' Side effects: tblSettings rows and one "Change ingest settings" audit row,
'   in one transaction.
Public Function IngestSaveSettings(ByVal dropFolder As String, ByVal mailbox As String, ByVal mailFolder As String, _
        ByVal defClass As String, ByVal maxClass As String, ByVal minutes As String) As String
    Dim ws As DAO.Workspace, inTrans As Boolean, keys As Variant, vals As Variant, i As Long, was As String, changes As String
    RequireAnchorAdmin
    dropFolder = Trim$(dropFolder): mailbox = Trim$(mailbox): mailFolder = Trim$(mailFolder)
    defClass = UCase$(Trim$(defClass)): maxClass = UCase$(Trim$(maxClass)): minutes = Trim$(minutes)
    If Right$(dropFolder, 1) = "\" Then dropFolder = Left$(dropFolder, Len(dropFolder) - 1)
    If Len(dropFolder) > 255 Or Len(mailbox) > 255 Or Len(mailFolder) > 255 Then
        IngestSaveSettings = "A folder or mailbox name can be at most 255 characters.": Exit Function
    End If
    If Len(dropFolder) > 0 And Not (Mid$(dropFolder, 2, 2) = ":\" Or Left$(dropFolder, 2) = "\\") Then
        IngestSaveSettings = "The drop folder must be a full path, such as D:\ANCHOR\ingest or \\server\share\ingest.": Exit Function
    End If
    If DCount("*", "tblMessageClasses", "ClassCode=" & Q(defClass)) = 0 Then
        IngestSaveSettings = "Pick the class received messages take (A to E).": Exit Function
    End If
    If InStr(1, ";UNCLASSIFIED;CUI;CONFIDENTIAL;SECRET;", ";" & maxClass & ";") = 0 Or Len(maxClass) = 0 Then
        IngestSaveSettings = "Pick the highest classification to take.": Exit Function
    End If
    If Len(minutes) = 0 Then minutes = "0"
    If Not minutes Like String$(Len(minutes), "#") Or Len(minutes) > 4 Then
        IngestSaveSettings = "Check every: a whole number of minutes, 0 to turn automatic checks off.": Exit Function
    End If
    If CLng(minutes) > 1440 Then IngestSaveSettings = "Check every: at most 1440 minutes (a day).": Exit Function
    minutes = CStr(CLng(minutes))
    keys = Array(INGEST_SET_FOLDER, INGEST_SET_MAILBOX, INGEST_SET_MAILFOLDER, INGEST_SET_CLASS, INGEST_SET_MAX, INGEST_SET_MINUTES)
    vals = Array(dropFolder, mailbox, mailFolder, defClass, maxClass, minutes)
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    ws.BeginTrans: inTrans = True
    For i = 0 To UBound(keys)
        was = Nz(DLookup("SettingValue", "tblSettings", "SettingKey=" & Q(keys(i))), "")
        If StrComp(was, vals(i), vbBinaryCompare) <> 0 Then
            IngPutSetting CStr(keys(i)), CStr(vals(i))
            changes = changes & IIf(Len(changes) > 0, "; ", "") & keys(i) & " was """ & was & """, now """ & vals(i) & """"
        End If
    Next
    If Len(changes) > 0 Then AuditLog "Change ingest settings", "Settings (mail ingest)", changes & "."
    ws.CommitTrans: inTrans = False
    Exit Function
Fail:
    If inTrans Then RollbackKeepingError ws
    RaiseAgain "modAnchorIngest.IngestSaveSettings"
End Function

' Set a tblSettings value, adding the row when there is none ("" is Null).
Private Sub IngPutSetting(ByVal key As String, ByVal v As String)
    If KitExecute("UPDATE tblSettings SET SettingValue=" & IngSqlText(v) & " WHERE SettingKey=" & Q(key)) = 0 Then
        KitExecute "INSERT INTO tblSettings (SettingKey, SettingValue, Description) VALUES (" & Q(key) & ", " & IngSqlText(v) & _
            ", 'Mail ingest')"
    End If
End Sub

' ---------------------------------------------------------------------------
' The feeders
' ---------------------------------------------------------------------------

' Ingest every .txt file in the drop folder (UTF-8 or ANSI), moving each to
' its Ingested or Rejected subfolder (made when missing); a rejected one
' gets a .reason.txt beside it. A file that fails for another reason (a
' lock, a read error) stays where it is for the next run and is counted as
' an error.
'   problem   set to why the folder couldn't be checked ("" when it was)
' Returns the run's summary (IngestLastRunText).
' Errors: ERR_NOT_ALLOWED unless a System Admin now (also mid-run).
' Make the drop folder and its Ingested and Rejected subfolders when
' missing. Returns "" or why they can't be used. A real error is caught by a
' handler rather than read from Err after On Error Resume Next: the kit's
' FolderExists leaves Err set (53, File not found) for a folder that isn't
' there yet, which made the first run on a new folder look like a failure.
Private Function IngMakeFolders(ByVal folder As String) As String
    On Error GoTo Fail
    EnsureFolder folder
    EnsureFolder folder & "\Ingested"
    EnsureFolder folder & "\Rejected"
    Exit Function
Fail:
    IngMakeFolders = "The drop folder " & folder & " can't be used: " & Err.Description
End Function

Public Function IngestCheckFolder(ByRef problem As String) As String
    Dim folder As String, f As String, files As Collection, v As Variant, n As Long
    problem = ""
    RequireAnchorAdmin
    If Not IngestReady() Then problem = "This data file has no ingest log. Make the data file again to add mail ingestion.": Exit Function
    folder = IngestDropFolder()
    If Len(folder) = 0 Then problem = "Set the drop folder first.": Exit Function
    problem = IngMakeFolders(folder)
    If Len(problem) > 0 Then Exit Function
    On Error GoTo Fail
    IngStartRun "Folder"
    Set files = New Collection
    f = Dir$(folder & "\*.txt", vbNormal)
    Do While Len(f) > 0
        If LCase$(Right$(f, 4)) = ".txt" Then files.Add f      ' Dir also matches "*.txtx" by its short name
        f = Dir$()
    Loop
    For Each v In files
        n = n + 1
        SysCmd acSysCmdSetStatus, "Ingesting " & n & " of " & files.Count & ": " & v
        IngOneFile folder, CStr(v)
    Next
    SysCmd acSysCmdClearStatus
    IngestCheckFolder = IngFinishRun()
    Exit Function
Fail:
    IngClearStatusKeepingError
    RaiseAgain "modAnchorIngest.IngestCheckFolder", folder
End Function

' Ingest one file of the drop folder and move it.
Private Sub IngOneFile(ByVal folder As String, ByVal fileName As String)
    Dim path As String, raw As String, res As String, dest As String
    On Error GoTo Fail
    path = folder & "\" & fileName
    raw = IngReadText(path, READ_LIMIT)
    res = IngestRaw(raw, "Folder", "", fileName)
    IngCount res
    If res = INGEST_STORED Or res = INGEST_REPEAT Then
        IngMoveFile path, folder & "\Ingested"
    Else
        dest = IngMoveFile(path, folder & "\Rejected")
        IngWriteReason dest, res, m_lastReason
    End If
    Exit Sub
Fail:
    If Err.Number = ERR_NOT_ALLOWED Then RaiseAgain "modAnchorIngest.IngOneFile", fileName
    If Len(m_runError) = 0 Then m_runError = fileName & ": " & Err.Description
    m_runErrors = m_runErrors + 1
    LogError "modAnchorIngest.IngOneFile", fileName
End Sub

' Move a file into folder dest, as "name (2).txt" when the name is taken
' there. Returns its new path.
Private Function IngMoveFile(ByVal path As String, ByVal dest As String) As String
    Dim nm As String, base As String, ext As String, target As String, n As Long
    nm = Mid$(path, InStrRev(path, "\") + 1)
    If InStrRev(nm, ".") > 0 Then
        base = Left$(nm, InStrRev(nm, ".") - 1): ext = Mid$(nm, InStrRev(nm, "."))
    Else
        base = nm
    End If
    target = dest & "\" & nm
    n = 1
    Do While FileExists(target)
        n = n + 1
        target = dest & "\" & base & " (" & n & ")" & ext
    Loop
    Name path As target
    IngMoveFile = target
End Function

' Write "<name>.reason.txt" beside a rejected file.
Private Sub IngWriteReason(ByVal path As String, ByVal res As String, ByVal reason As String)
    Dim ff As Integer, rp As String
    If LCase$(Right$(path, 4)) = ".txt" Then rp = Left$(path, Len(path) - 4) & ".reason.txt" Else rp = path & ".reason.txt"
    ff = FreeFile
    Open rp For Output As #ff
    Print #ff, "ANCHOR mail ingest, " & Format$(Now(), "dd mmm yyyy hh:nn") & ", " & SignedInName()
    Print #ff, "Result: " & res
    Print #ff, "Reason: " & reason
    Close #ff
End Sub

' A text file's text, read as UTF-8 (with or without its byte-order mark),
' UTF-16 (with one) or the computer's ANSI code page; at most maxBytes are
' read.
Private Function IngReadText(ByVal path As String, ByVal maxBytes As Long) As String
    Dim ff As Integer, n As Long, b() As Byte, w() As Byte, i As Long
    ff = FreeFile
    Open path For Binary Access Read Shared As #ff
    n = LOF(ff)
    If n > maxBytes Then n = maxBytes
    If n = 0 Then Close #ff: Exit Function
    ReDim b(0 To n - 1)
    Get #ff, , b
    Close #ff
    If n >= 2 Then
        If b(0) = &HFF And b(1) = &HFE Then
            If n > 2 Then
                ReDim w(0 To n - 3)
                For i = 2 To n - 1
                    w(i - 2) = b(i)
                Next
                IngReadText = w
            End If
            Exit Function
        End If
    End If
    If IngIsUtf8(b) Then
        IngReadText = IngUtf8Text(b)
    Else
        IngReadText = StrConv(b, vbUnicode)
    End If
End Function

' True when the bytes are valid UTF-8 (plain ASCII is too).
Private Function IngIsUtf8(b() As Byte) As Boolean
    Dim i As Long, n As Long, k As Long, need As Long
    n = UBound(b)
    i = 0
    Do While i <= n
        If b(i) < &H80 Then
            need = 0
        ElseIf (b(i) And &HE0) = &HC0 Then
            need = 1
        ElseIf (b(i) And &HF0) = &HE0 Then
            need = 2
        ElseIf (b(i) And &HF8) = &HF0 Then
            need = 3
        Else
            Exit Function
        End If
        For k = 1 To need
            If i + k > n Then Exit Function
            If (b(i + k) And &HC0) <> &H80 Then Exit Function
        Next
        i = i + need + 1
    Loop
    IngIsUtf8 = True
End Function

' UTF-8 bytes as text (a byte-order mark is dropped).
Private Function IngUtf8Text(b() As Byte) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1                                 ' binary
    st.Open
    st.Write b
    st.Position = 0
    st.Type = 2                                 ' text
    st.Charset = "utf-8"
    IngUtf8Text = st.ReadText(-1)
    st.Close
End Function

' Ingest the mail in the ingest mailbox's folder (IngestMailbox,
' IngestFolder) through Outlook: each item's .txt attachments, or its body
' when it has none, then the item to the folder's Ingested or Rejected
' subfolder (made when missing). The Internet Message-ID is logged when
' Outlook gives it.
'   problem   set to why the mailbox couldn't be checked (no Outlook, no
'             such mailbox or folder); "" when it was
' Returns the run's summary (IngestLastRunText).
' Errors: ERR_NOT_ALLOWED unless a System Admin now (also mid-run).
Public Function IngestCheckMailbox(ByRef problem As String) As String
    Dim ol As Object, ns As Object, src As Object, okF As Object, badF As Object, its As Object, it As Object
    Dim mailbox As String, fldPath As String, i As Long, n As Long, cls As Long
    problem = ""
    RequireAnchorAdmin
    If Not IngestReady() Then problem = "This data file has no ingest log. Make the data file again to add mail ingestion.": Exit Function
    mailbox = Trim$(SettingText(INGEST_SET_MAILBOX, ""))
    fldPath = Trim$(SettingText(INGEST_SET_MAILFOLDER, ""))
    If Len(fldPath) = 0 Then fldPath = "Inbox"
    If Len(mailbox) = 0 Then problem = "Set the mailbox first: its name as Outlook shows it, or its address.": Exit Function
    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    If ol Is Nothing Then
        problem = "Outlook isn't available on this computer, so the mailbox can't be checked from here (" & Err.Description & _
            "). Check it from a computer with Outlook and a mail profile, or use the drop folder."
        Exit Function
    End If
    Set ns = ol.GetNamespace("MAPI")
    If ns Is Nothing Then problem = "Outlook has no mail profile on this computer (" & Err.Description & ").": Exit Function
    Err.Clear
    On Error GoTo Fail
    Set src = IngMailFolder(ns, mailbox, fldPath, problem)
    If src Is Nothing Then Exit Function
    Set okF = IngSubFolder(src, "Ingested")
    Set badF = IngSubFolder(src, "Rejected")
    IngStartRun "Outlook"
    Set its = src.Items
    ' Backwards: each item moves out of the folder as it is done.
    For i = its.Count To 1 Step -1
        Set it = its.Item(i)
        cls = 0
        On Error Resume Next
        cls = it.Class
        On Error GoTo Fail
        If cls = OL_MAIL_ITEM Then
            n = n + 1
            SysCmd acSysCmdSetStatus, "Ingesting mail item " & n
            IngOneItem it, okF, badF
        End If
    Next
    SysCmd acSysCmdClearStatus
    IngestCheckMailbox = IngFinishRun()
    Exit Function
Fail:
    IngClearStatusKeepingError
    RaiseAgain "modAnchorIngest.IngestCheckMailbox", mailbox & " / " & fldPath
End Function

' The Outlook folder fldPath ("Inbox", "Inbox\ANCHOR") of mailbox, a store
' in the profile by its name, or else a shared mailbox by its address
' (starting at its Inbox). Nothing with why when it isn't there.
Private Function IngMailFolder(ByVal ns As Object, ByVal mailbox As String, ByVal fldPath As String, ByRef why As String) As Object
    Dim cur As Object, nxt As Object, rcp As Object, parts() As String, i As Long, first As Long
    parts = Split(Replace(fldPath, "/", "\"), "\")
    On Error Resume Next
    Set cur = ns.Folders.Item(mailbox)
    On Error GoTo 0
    If cur Is Nothing Then
        On Error Resume Next
        Set rcp = ns.CreateRecipient(mailbox)
        rcp.Resolve
        If rcp.Resolved Then Set cur = ns.GetSharedDefaultFolder(rcp, OL_FOLDER_INBOX)
        On Error GoTo 0
        If cur Is Nothing Then
            why = "Outlook on this computer has no mailbox """ & mailbox & """. Use its name as Outlook's folder list shows it, " & _
                "or its address."
            Exit Function
        End If
        If StrComp(Trim$(parts(0)), "Inbox", vbTextCompare) = 0 Then first = 1
    End If
    For i = first To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            Set nxt = Nothing
            On Error Resume Next
            Set nxt = cur.Folders.Item(Trim$(parts(i)))
            On Error GoTo 0
            If nxt Is Nothing Then why = "The mailbox """ & mailbox & """ has no folder """ & fldPath & """.": Exit Function
            Set cur = nxt
        End If
    Next
    Set IngMailFolder = cur
End Function

' The subfolder nm of an Outlook folder, made when it is missing.
Private Function IngSubFolder(ByVal parent As Object, ByVal nm As String) As Object
    Dim f As Object
    On Error Resume Next
    Set f = parent.Folders.Item(nm)
    On Error GoTo 0
    If f Is Nothing Then Set f = parent.Folders.Add(nm)
    Set IngSubFolder = f
End Function

' Ingest one mail item and move it: to Rejected when any text in it was
' refused, else to Ingested.
Private Sub IngOneItem(ByVal it As Object, ByVal okF As Object, ByVal badF As Object)
    Dim att As Object, msgId As String, label As String, tmp As String, raw As String, res As String, n As Long, bad As Boolean
    Dim num As Long, desc As String, src As String
    On Error GoTo Fail
    msgId = IngInternetId(it)
    label = "Mail: " & ClipText(Nz(it.Subject, "(no subject)"), 120)
    On Error Resume Next
    label = label & " (" & Format$(it.ReceivedTime, "dd mmm yyyy hh:nn") & ")"
    On Error GoTo Fail
    For Each att In it.Attachments
        If LCase$(Right$(att.FileName, 4)) = ".txt" Then
            n = n + 1
            tmp = Environ$("TEMP") & "\anchor-ingest-" & Format$(Now(), "yyyymmddhhnnss") & "-" & n & ".txt"
            att.SaveAsFile tmp
            raw = IngReadText(tmp, READ_LIMIT)
            Kill tmp                            ' this run's own copy; the mail item keeps the attachment
            tmp = ""
            res = IngestRaw(raw, "Outlook", msgId, ClipText(label & " / " & att.FileName, 255))
            IngCount res
            If res = INGEST_REJECTED Or res = INGEST_CONFLICT Then bad = True
        End If
    Next
    If n = 0 Then
        res = IngestRaw(Nz(it.Body, ""), "Outlook", msgId, label)
        IngCount res
        If res = INGEST_REJECTED Or res = INGEST_CONFLICT Then bad = True
    End If
    If bad Then it.Move badF Else it.Move okF
    Exit Sub
Fail:
    ' The temp copy goes before anything else; the error is kept (an On
    ' Error statement would clear it).
    num = Err.Number: desc = Err.Description: src = Err.Source
    If Len(tmp) > 0 Then
        On Error Resume Next
        Kill tmp
    End If
    On Error GoTo 0
    If num = ERR_NOT_ALLOWED Then Err.Raise num, src, desc
    m_runErrors = m_runErrors + 1
    If Len(m_runError) = 0 Then m_runError = label & ": " & desc
    Err.Number = num: Err.Description = desc: Err.Source = src
    LogError "modAnchorIngest.IngOneItem", label
End Sub

' Clear the status bar message from an error handler, leaving the error
' being handled as it was.
Private Sub IngClearStatusKeepingError()
    Dim num As Long, desc As String, src As String
    num = Err.Number: desc = Err.Description: src = Err.Source
    On Error Resume Next
    SysCmd acSysCmdClearStatus
    On Error GoTo 0
    Err.Number = num: Err.Description = desc: Err.Source = src
End Sub

' The mail's Internet Message-ID (PR_INTERNET_MESSAGE_ID), "" when Outlook
' doesn't give it.
Private Function IngInternetId(ByVal it As Object) As String
    Dim v As Variant
    On Error Resume Next
    v = it.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID)
    If Err.Number = 0 Then IngInternetId = Left$(Trim$(Nz(v, "")), 255)
End Function

' Start counting a run.
Private Sub IngStartRun(ByVal feeder As String)
    m_runFeeder = feeder
    m_runStored = 0: m_runRepeat = 0: m_runConflict = 0: m_runRejected = 0: m_runErrors = 0: m_runError = ""
End Sub

Private Sub IngCount(ByVal res As String)
    Select Case res
        Case INGEST_STORED: m_runStored = m_runStored + 1
        Case INGEST_REPEAT: m_runRepeat = m_runRepeat + 1
        Case INGEST_CONFLICT: m_runConflict = m_runConflict + 1
        Case Else: m_runRejected = m_runRejected + 1
    End Select
End Sub

' The run's summary, kept in tblSettings IngestLastRun. Returns it.
Private Function IngFinishRun() As String
    Dim s As String
    s = Format$(Now(), "dd mmm yyyy hh:nn") & ", " & m_runFeeder & ": "
    If m_runStored + m_runRepeat + m_runConflict + m_runRejected + m_runErrors = 0 Then
        s = s & "nothing waiting."
    Else
        s = s & m_runStored & " stored, " & m_runRepeat & " repeat, " & m_runConflict & " conflict, " & m_runRejected & " rejected"
        If m_runErrors > 0 Then s = s & ", " & m_runErrors & " " & Plural(m_runErrors, "error") & " (left in place; first: " & ClipText(m_runError, 120) & ")"
        s = s & "."
    End If
    IngPutSetting INGEST_SET_LASTRUN, Left$(s, 255)
    IngFinishRun = s
End Function

' ---------------------------------------------------------------------------
' The Admin page's list
' ---------------------------------------------------------------------------

' The ingest log, newest first (the latest 500), for the Admin page's
' Mail ingest tab: id, when, feeder, serial, result, and the reason with
' the file or mail item.
Public Function IngestLogSql() As String
    IngestLogSql = "SELECT TOP 500 l.IngestID AS id, l.LoggedAt, Format(l.LoggedAt,'dd mmm hh:nn') AS logged_txt, l.Feeder, " & _
        "l.Reference, UCase(Left(l.Result,1)) & Mid(l.Result,2) AS result_txt, " & _
        "Nz(l.Reason,'') & IIf(Len(Nz(l.Reason,''))>0 And Len(Nz(l.FileOrItem,''))>0,' | ','') & Nz(l.FileOrItem,'') AS detail " & _
        "FROM tblIngestLog AS l ORDER BY l.LoggedAt DESC, l.IngestID DESC"
End Function
