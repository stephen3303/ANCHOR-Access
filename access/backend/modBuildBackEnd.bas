Attribute VB_Name = "modBuildBackEnd"
Option Compare Database
Option Explicit
' ANCHOR - back-end builder.
' Creates the shared data file (ANCHOR_BE.accdb) with every table, index and
' relationship, the lookup lists, and the web app's fictional demo traffic
' from backend\seed\*.tsv (tools\make_seed.py writes them). Two ways in:
'   BuildBackEnd                  run by hand from any Access database; asks for
'                                 the path and the seed folder.
'   BuildBackEndAt(path, seed)    run by backend\new-backend.ps1; asks nothing,
'                                 never overwrites, and returns "OK" or what
'                                 failed. seed "" makes an empty data file.
'
' People: names, accounts and email live once in the kit's shared Personnel file
' (kit_person). This file keeps only each person's ANCHOR profile (billet, unit,
' community) and roles, keyed on the Personnel person_id (tblPeople,
' tblPersonRoles). No people are seeded: the first Personnel admin to open the
' app while tblPeople is empty becomes its admin, once (modSignIn). Seeded rows
' that name a person (a subscription request, an audit row) hold the web
' data's name as text and no person_id.
'
' Messages: tblMessages holds each message as it arrived (RawText, verbatim)
' and as ANCHOR formatted it (BodyText, and its sections in
' tblMessageSections), with the integrity data recorded at ingestion (C2OIX
' id, SHA-256, size). The AI summary (AiSummary) is stored text, never the
' official record; the app shows it apart, labelled "Not authoritative".
'
' Audit: tblAudit is append-only. Each row's RowHash is the SHA-256 of the row
' and the previous row's hash (modAudit in the front end), so a changed or
' removed row breaks the chain. This builder has no SHA-256, so the seeded
' rows go in with no hash; the first front end to open the data file seals
' them in order (modAudit.AuditSealPending).
'
' help_topic holds the help text the admin edits in the app; its columns match
' the kit's zz_help. error_log is the kit's error log (modErrors).
'
' The schema version (tblSettings SchemaVersion) is the app's DATA_VERSION
' (modSignIn). A change here bumps both and adds a step to modSignIn.UpgradeData.

' Make the data file by hand from any Access database.
Public Sub BuildBackEnd()
    Dim p As String, seed As String, r As String
    p = InputBox("Back-end path:", "ANCHOR", CurrentProject.Path & "\ANCHOR_BE.accdb")
    If Len(p) = 0 Then Exit Sub
    seed = InputBox("Seed folder (blank for an empty data file):", "ANCHOR", CurrentProject.Path & "\seed")
    If Len(Dir(p)) > 0 Then
        If MsgBox(p & " exists. Overwrite?", vbYesNo + vbExclamation) <> vbYes Then Exit Sub
        Kill p
    End If
    r = BuildBackEndAt(p, seed)
    If r = "OK" Then MsgBox "Back end created:" & vbCrLf & p, vbInformation Else MsgBox r, vbCritical
End Sub

' Script entry point: create the data file at p (a full path to a new .accdb)
' and load the seed rows from folder seed ("" for none). Returns "OK", or
' "ERROR: <description>" with nothing left behind.
Public Function BuildBackEndAt(ByVal p As String, ByVal seed As String) As String
    Dim db As DAO.Database, stepName As String
    On Error GoTo Fail
    If Len(Dir(p)) > 0 Then Err.Raise 58, , "There is already a file at " & p & "; nothing was changed."
    Set db = DBEngine.CreateDatabase(p, dbLangGeneral, dbVersion120)
    stepName = "tables": CreateTables db
    stepName = "indexes": CreateIndexes db
    stepName = "relationships": CreateRelationships db
    stepName = "lookups": SeedLookups db
    If Len(seed) > 0 Then stepName = "seed rows": SeedDemo db, seed
    db.Close: Set db = Nothing
    BuildBackEndAt = "OK"
    Exit Function
Fail:
    BuildBackEndAt = "ERROR during " & stepName & ": " & Err.Description
    If Not db Is Nothing Then
        db.Close: Set db = Nothing
        On Error Resume Next
        Kill p                                  ' don't leave a half-made file
    End If
End Function

Private Sub ExecDdl(db As DAO.Database, ByVal sql As String)
    db.Execute sql, dbFailOnError
End Sub

Private Sub CreateTables(db As DAO.Database)
    ' ---- Settings, roles and people (keyed on the Personnel file's person_id) ----
    ExecDdl db, "CREATE TABLE tblSettings (SettingKey TEXT(50) PRIMARY KEY, SettingValue TEXT(255), Description TEXT(255))"
    ExecDdl db, "CREATE TABLE tblRoles (RoleID AUTOINCREMENT PRIMARY KEY, RoleCode TEXT(20) NOT NULL, RoleName TEXT(50) NOT NULL, " & _
      "Detail TEXT(255), SortOrder INTEGER)"
    ExecDdl db, "CREATE TABLE tblPeople (PersonID LONG PRIMARY KEY, Billet TEXT(100), Unit TEXT(60), Community TEXT(60), " & _
      "Location TEXT(60), Dsn TEXT(30), Groups TEXT(255), IsAdmin YESNO, IsActive YESNO, NotifyEmail YESNO, LastSeenOn DATETIME)"
    ExecDdl db, "CREATE TABLE tblPersonRoles (PersonID LONG NOT NULL, RoleID LONG NOT NULL, CONSTRAINT pkPersonRoles PRIMARY KEY (PersonID, RoleID))"

    ' ---- Message types and communications classes (Stephen's list, 2026-11-05) ----
    ExecDdl db, "CREATE TABLE tblMessageTypes (TypeID AUTOINCREMENT PRIMARY KEY, TypeCode TEXT(30) NOT NULL, TypeName TEXT(100), " & _
      "Category TEXT(30), Description TEXT(255), IsActive YESNO, SortOrder INTEGER)"
    ExecDdl db, "CREATE TABLE tblMessageClasses (ClassCode TEXT(1) PRIMARY KEY, ClassName TEXT(100), Description TEXT(255))"

    ' ---- Commands and address groups (ack tracking expands a group to its commands) ----
    ExecDdl db, "CREATE TABLE tblAddressGroups (GroupID AUTOINCREMENT PRIMARY KEY, GroupName TEXT(100) NOT NULL)"
    ExecDdl db, "CREATE TABLE tblGroupMembers (GroupID LONG NOT NULL, CommandName TEXT(100) NOT NULL, " & _
      "CONSTRAINT pkGroupMembers PRIMARY KEY (GroupID, CommandName))"

    ' ---- Messages ----
    ExecDdl db, "CREATE TABLE tblMessages (MessageID AUTOINCREMENT PRIMARY KEY, MsgKey TEXT(40) NOT NULL, Dtg TEXT(20), " & _
      "ReceivedOn DATETIME NOT NULL, Subject TEXT(255) NOT NULL, Reference TEXT(60) NOT NULL, MsgType TEXT(30), " & _
      "Originator TEXT(60), Source TEXT(40), ToLine TEXT(255), InfoLine TEXT(255), Precedence TEXT(12) NOT NULL, " & _
      "Classification TEXT(20) NOT NULL, Importance TEXT(10), Lifecycle TEXT(12) NOT NULL, RelatesTo TEXT(60), " & _
      "ReplacedBy TEXT(60), Communities TEXT(255), AckRequired YESNO, AckDue DATETIME, ActionText TEXT(255), " & _
      "ActionSuspense TEXT(20), ActionAssignedTo TEXT(60), C2oixId TEXT(30), Sha256 TEXT(64), SizeKb LONG, " & _
      "MsgClass TEXT(1), Retention TEXT(40), Distribution TEXT(255), Related TEXT(255), AiSummary MEMO, RawText MEMO, BodyText MEMO)"
    ExecDdl db, "CREATE TABLE tblMessageSections (SectionID AUTOINCREMENT PRIMARY KEY, MessageID LONG NOT NULL, " & _
      "SortOrder INTEGER, Heading TEXT(100), Body MEMO)"
    ExecDdl db, "CREATE TABLE tblAttachments (AttachmentID AUTOINCREMENT PRIMARY KEY, MessageID LONG NOT NULL, " & _
      "FileName TEXT(150) NOT NULL, SizeText TEXT(20), Kind TEXT(10))"
    ExecDdl db, "CREATE TABLE tblAddressees (AddresseeID AUTOINCREMENT PRIMARY KEY, MessageID LONG NOT NULL, " & _
      "AddrName TEXT(100) NOT NULL, AddrRole TEXT(4) NOT NULL, AddrKind TEXT(10) NOT NULL)"
    ' Acknowledgements the originator received from addressee commands.
    ExecDdl db, "CREATE TABLE tblAckReceipts (ReceiptID AUTOINCREMENT PRIMARY KEY, MessageID LONG NOT NULL, " & _
      "CommandName TEXT(100) NOT NULL, ReceivedOn DATETIME NOT NULL)"
    ' Each person's state of each message: no row is unread.
    ExecDdl db, "CREATE TABLE tblMessageUserState (MessageID LONG NOT NULL, PersonID LONG NOT NULL, ReadState TEXT(10) NOT NULL, " & _
      "FirstReadOn DATETIME, AckOn DATETIME, CONSTRAINT pkMessageUserState PRIMARY KEY (MessageID, PersonID))"

    ' ---- Subscriptions: the catalogue, and each person's (or a request) ----
    ExecDdl db, "CREATE TABLE tblSubscriptions (SubID AUTOINCREMENT PRIMARY KEY, SubKey TEXT(40), SubName TEXT(100) NOT NULL, " & _
      "Category TEXT(60), Description TEXT(255), MsgTypes TEXT(255), Originators TEXT(255), Keywords TEXT(255), " & _
      "Precedences TEXT(100), IsMandatory YESNO, DefaultFrequency TEXT(20), IsActive YESNO)"
    ExecDdl db, "CREATE TABLE tblPersonSubs (PersonSubID AUTOINCREMENT PRIMARY KEY, SubID LONG NOT NULL, PersonID LONG, " & _
      "PersonName TEXT(100) NOT NULL, Status TEXT(20) NOT NULL, Frequency TEXT(20), ChInApp YESNO, ChEmail YESNO, ChTeams YESNO, " & _
      "Justification MEMO, RequestedOn DATETIME, LastUpdated DATETIME, DeciderID LONG, DeciderName TEXT(100), " & _
      "DecidedOn DATETIME, DecisionNote TEXT(255))"

    ' ---- Mandatory alert rules and command access groups (Admin) ----
    ExecDdl db, "CREATE TABLE tblAlertRules (RuleID AUTOINCREMENT PRIMARY KEY, RuleKind TEXT(40) NOT NULL, MatchText TEXT(100) NOT NULL, " & _
      "Audience TEXT(100), AckHours INTEGER, Escalate TEXT(100), IsActive YESNO, SortOrder INTEGER)"
    ExecDdl db, "CREATE TABLE tblAccessGroups (AccessGroupID AUTOINCREMENT PRIMARY KEY, GroupName TEXT(100) NOT NULL, " & _
      "Basis TEXT(40), Members LONG, Scope TEXT(150))"

    ' ---- Audit trail (append-only, hash chained) ----
    ExecDdl db, "CREATE TABLE tblAudit (AuditID AUTOINCREMENT PRIMARY KEY, LoggedAt DATETIME NOT NULL, PersonID LONG, " & _
      "UserName TEXT(100) NOT NULL, Action TEXT(40) NOT NULL, Resource TEXT(150), Classification TEXT(20), Result TEXT(10), " & _
      "Device TEXT(64), Location TEXT(64), Details MEMO, PrevHash TEXT(64), RowHash TEXT(64))"

    ' ---- Search, briefings and the watch ----
    ExecDdl db, "CREATE TABLE tblSavedSearches (SearchID AUTOINCREMENT PRIMARY KEY, PersonID LONG NOT NULL, SearchName TEXT(100) NOT NULL, " & _
      "Criteria MEMO, CreatedOn DATETIME)"
    ExecDdl db, "CREATE TABLE tblBriefings (BriefingID AUTOINCREMENT PRIMARY KEY, Title TEXT(150) NOT NULL, PersonID LONG, " & _
      "PersonName TEXT(100), CreatedOn DATETIME, PeriodFrom DATETIME, PeriodTo DATETIME, Criteria MEMO, Notes MEMO)"
    ExecDdl db, "CREATE TABLE tblBriefingItems (BriefingID LONG NOT NULL, MessageID LONG NOT NULL, SortOrder INTEGER, ItemNote MEMO, " & _
      "CONSTRAINT pkBriefingItems PRIMARY KEY (BriefingID, MessageID))"
    ExecDdl db, "CREATE TABLE tblWatchNotes (NoteID AUTOINCREMENT PRIMARY KEY, NoteText MEMO NOT NULL, PersonID LONG, " & _
      "PersonName TEXT(100), CreatedOn DATETIME, IsOpen YESNO)"
    ExecDdl db, "CREATE TABLE tblOngoingTopics (TopicID AUTOINCREMENT PRIMARY KEY, Topic TEXT(100) NOT NULL, LastUpdate DATETIME, " & _
      "Status TEXT(20))"
    ExecDdl db, "CREATE TABLE tblTurnoverChecklist (ItemID AUTOINCREMENT PRIMARY KEY, ItemText TEXT(150) NOT NULL, SortOrder INTEGER, IsActive YESNO)"
    ExecDdl db, "CREATE TABLE tblTurnovers (TurnoverID AUTOINCREMENT PRIMARY KEY, PersonID LONG, PersonName TEXT(100), " & _
      "StartedOn DATETIME, CompletedOn DATETIME, Notes MEMO)"
    ExecDdl db, "CREATE TABLE tblTurnoverChecks (TurnoverID LONG NOT NULL, ItemID LONG NOT NULL, CheckedOn DATETIME, " & _
      "CONSTRAINT pkTurnoverChecks PRIMARY KEY (TurnoverID, ItemID))"

    ' ---- Help text (admin-editable; same columns as the kit's zz_help) ----
    ExecDdl db, "CREATE TABLE help_topic (topic_key TEXT(32) CONSTRAINT pk_help_topic PRIMARY KEY, title TEXT(80), sort_order LONG, " & _
      "body MEMO, search_text MEMO, updated_at DATETIME, updated_by TEXT(150))"

    ' ---- Error log (modErrors writes it) ----
    ExecDdl db, "CREATE TABLE error_log (id COUNTER CONSTRAINT pk_error_log PRIMARY KEY, logged_at DATETIME, " & _
      "error_number LONG, description LONGTEXT, err_source TEXT(255), source TEXT(255), call_trail LONGTEXT, " & _
      "context LONGTEXT, page TEXT(32), active_form TEXT(64), active_control TEXT(64), username TEXT(150), " & _
      "windows_user TEXT(64), computer TEXT(64), fe_build TEXT(32), access_version TEXT(32))"
    CreateDraftingTables db         ' drafting and release (the block at the end)
    CreateIngestTables db           ' mail ingestion (the last block)
End Sub

Private Sub CreateIndexes(db As DAO.Database)
    ExecDdl db, "CREATE UNIQUE INDEX uxRoleCode ON tblRoles (RoleCode)"
    ExecDdl db, "CREATE UNIQUE INDEX uxTypeCode ON tblMessageTypes (TypeCode)"
    ExecDdl db, "CREATE UNIQUE INDEX uxGroupName ON tblAddressGroups (GroupName)"
    ExecDdl db, "CREATE UNIQUE INDEX uxMsgKey ON tblMessages (MsgKey)"
    ExecDdl db, "CREATE INDEX ixMsgReceived ON tblMessages (ReceivedOn)"
    ExecDdl db, "CREATE INDEX ixMsgReference ON tblMessages (Reference)"
    ExecDdl db, "CREATE INDEX ixSectionMsg ON tblMessageSections (MessageID, SortOrder)"
    ExecDdl db, "CREATE INDEX ixAttachMsg ON tblAttachments (MessageID)"
    ExecDdl db, "CREATE INDEX ixAddrMsg ON tblAddressees (MessageID)"
    ExecDdl db, "CREATE INDEX ixReceiptMsg ON tblAckReceipts (MessageID)"
    ExecDdl db, "CREATE INDEX ixStatePerson ON tblMessageUserState (PersonID)"
    ExecDdl db, "CREATE UNIQUE INDEX uxSubName ON tblSubscriptions (SubName)"
    ExecDdl db, "CREATE INDEX ixPersonSubs ON tblPersonSubs (PersonID, SubID)"
    ExecDdl db, "CREATE INDEX ixPersonSubStatus ON tblPersonSubs (Status)"
    ExecDdl db, "CREATE INDEX ixAuditAt ON tblAudit (LoggedAt)"
    ExecDdl db, "CREATE INDEX ixSavedSearchPerson ON tblSavedSearches (PersonID)"
    ExecDdl db, "CREATE INDEX ixErrorLogged ON error_log (logged_at)"
    CreateDraftingIndexes db
    CreateIngestIndexes db
End Sub

' Relationship nm: ct.fk refers to pt.pk, enforced, deleting the children
' with the parent when cascadeDel.
Private Sub AddRelation(db As DAO.Database, ByVal nm As String, ByVal pt As String, ByVal pk As String, ByVal ct As String, _
        ByVal fk As String, Optional ByVal cascadeDel As Boolean = False)
    Dim r As DAO.Relation
    Set r = db.CreateRelation(nm, pt, ct, IIf(cascadeDel, dbRelationDeleteCascade, 0))
    r.Fields.Append r.CreateField(pk)
    r.Fields(pk).ForeignName = fk
    db.Relations.Append r
End Sub

Private Sub CreateRelationships(db As DAO.Database)
    AddRelation db, "rPR_P", "tblPeople", "PersonID", "tblPersonRoles", "PersonID", True
    AddRelation db, "rPR_R", "tblRoles", "RoleID", "tblPersonRoles", "RoleID"
    AddRelation db, "rGM_G", "tblAddressGroups", "GroupID", "tblGroupMembers", "GroupID", True
    AddRelation db, "rMS_M", "tblMessages", "MessageID", "tblMessageSections", "MessageID", True
    AddRelation db, "rAt_M", "tblMessages", "MessageID", "tblAttachments", "MessageID", True
    AddRelation db, "rAd_M", "tblMessages", "MessageID", "tblAddressees", "MessageID", True
    AddRelation db, "rAR_M", "tblMessages", "MessageID", "tblAckReceipts", "MessageID", True
    AddRelation db, "rUS_M", "tblMessages", "MessageID", "tblMessageUserState", "MessageID", True
    AddRelation db, "rPS_S", "tblSubscriptions", "SubID", "tblPersonSubs", "SubID"
    AddRelation db, "rBI_B", "tblBriefings", "BriefingID", "tblBriefingItems", "BriefingID", True
    AddRelation db, "rBI_M", "tblMessages", "MessageID", "tblBriefingItems", "MessageID"
    AddRelation db, "rTC_T", "tblTurnovers", "TurnoverID", "tblTurnoverChecks", "TurnoverID", True
    AddRelation db, "rTC_I", "tblTurnoverChecklist", "ItemID", "tblTurnoverChecks", "ItemID"
    CreateDraftingRelationships db
End Sub

' The lists every data file has, demo traffic or not.
Private Sub SeedLookups(db As DAO.Database)
    ExecDdl db, "INSERT INTO tblSettings VALUES ('SchemaVersion','1','Back-end schema version')"
    ExecDdl db, "INSERT INTO tblSettings VALUES ('CommandSource','COMSUBLANT','The source code of this command''s own traffic')"
    ExecDdl db, "INSERT INTO tblSettings VALUES ('CommandName','Commander, Submarine Force Atlantic','This command''s name')"
    SeedMessageTypes db
    ' Business rules that make delivery mandatory (the web app's Admin page).
    AddRule db, 1, "Precedence", "FLASH, IMMEDIATE", "All addressed units", 2, "Department Head, then XO"
    AddRule db, 2, "Message type", "SAFETY MSG", "All personnel", 24, "Safety Officer"
    AddRule db, 3, "Command applicability", "COMSUBLANT action messages", "Commanding Officers, XOs", 48, "SUBRON Commander"
    AddRule db, 4, "Role", "Cybersecurity directives", "ISSM, IT Chief", 72, "Comms Officer"
    AddAccess db, "COMSUBLANT", "Organizational affiliation", 412, "All COMSUBLANT traffic"
    AddAccess db, "COMSUBLANT N6", "Role", 14, "Directorate traffic, officer-only"
    AddAccess db, "Comms Division", "Role", 9, "Unit traffic, comms admin"
    AddAccess db, "SUBRON Staff", "Group", 38, "Squadron traffic"
    SeedDraftingLookups db
    SeedIngestSettings db
End Sub

' The message types (administrative, operational, safety, and the web app's
' command traffic types) and the five communications classes. Admins can add
' more; drafting and search pick from this list.
Private Sub SeedMessageTypes(db As DAO.Database)
    Dim t As Variant, p() As String, i As Long
    For Each t In Array( _
            "NAVADMIN|Navy administrative message|Administrative & General|Navy-specific administrative messages to all Navy units and the Marine Corps", _
            "ALNAV|All Navy message|Administrative & General|SECNAV messages to all Navy and Marine Corps activities", _
            "ALNAVRESFOR|All Navy Reserve Force message|Administrative & General|Navy Reserve administrative messages", _
            "ALSAFE|All safety message|Administrative & General|Naval Safety Command safety-related messages", _
            "GENADMIN|General administrative message|Administrative & General|Standard narrative format for most administrative messages", _
            "CIM|Communications Improvement Memorandum|Administrative & General|Post-incident training notice on procedural errors", _
            "OPGEN|Operational general message|Operational|Policy, general instructions and common warfare information", _
            "OPTASK|Operational task message|Operational|Specific tasking and resource allocation", _
            "OPSTAT|Operational status message|Operational|Status reports and operational data", _
            "WARNORD|Warning Order|Operational|Preliminary notice of impending action, to allow planning", _
            "LOI|Letter of Instruction|Operational|Planning guidance when a formal order is not required", _
            "OPNAVINST|OPNAV instruction|Web app traffic|Chief of Naval Operations instruction", _
            "MARADMIN|Marine Corps administrative message|Web app traffic|Marine Corps administrative messages", _
            "NAVSEAINST|NAVSEA instruction|Web app traffic|Naval Sea Systems Command instruction", _
            "COMSUBLANT MSG|COMSUBLANT message|Web app traffic|Commander, Submarine Force Atlantic traffic", _
            "COMSUBFOR INST|COMSUBFOR instruction|Web app traffic|Submarine Force instruction", _
            "FLEET MSG|Fleet message|Web app traffic|Fleet-wide operational traffic", _
            "SAFETY MSG|Safety message|Web app traffic|Safety direction to all personnel", _
            "HYDRO|Hydrographic advisory|Special|Hydrographic and navigation advisories", _
            "NOTAM|Notice to Airmen and Mariners|Special|Navigation advisories", _
            "Q MESSAGE|Q message|Special|Specialized question and answer messages", _
            "ALL SHIPS PRESENT|All ships present message|Special|Broadcast to all ships in a designated area", _
            "RED CROSS|Red Cross message|Special|Humanitarian and Red Cross communications")
        p = Split(t, "|"): i = i + 1
        ExecDdl db, "INSERT INTO tblMessageTypes (TypeCode, TypeName, Category, Description, IsActive, SortOrder) VALUES ('" & _
            p(0) & "','" & p(1) & "','" & p(2) & "','" & p(3) & "',True," & i & ")"
    Next
    For Each t In Array("A|Official Department of Defense|Official DoD messages (the largest volume)", _
            "B|Other U.S. Government|Other U.S. Government departments and agencies", _
            "C|Broadcast|Broadcast traffic for all nationalities (weather, hydrographic data)", _
            "D|Private, toll|Private messages for which the Navy collects tolls (press)", _
            "E|Personal|Personal messages to and from naval personnel, free over Navy circuits")
        p = Split(t, "|")
        ExecDdl db, "INSERT INTO tblMessageClasses (ClassCode, ClassName, Description) VALUES ('" & p(0) & "','" & p(1) & "','" & p(2) & "')"
    Next
End Sub

Private Sub AddRule(db As DAO.Database, ByVal ord As Long, ByVal kind As String, ByVal match As String, ByVal aud As String, _
        ByVal hrs As Long, ByVal esc As String)
    ExecDdl db, "INSERT INTO tblAlertRules (RuleKind, MatchText, Audience, AckHours, Escalate, IsActive, SortOrder) VALUES ('" & _
        kind & "','" & match & "','" & aud & "'," & hrs & ",'" & esc & "',True," & ord & ")"
End Sub

Private Sub AddAccess(db As DAO.Database, ByVal nm As String, ByVal basis As String, ByVal n As Long, ByVal scope As String)
    ExecDdl db, "INSERT INTO tblAccessGroups (GroupName, Basis, Members, Scope) VALUES ('" & nm & "','" & basis & "'," & n & ",'" & scope & "')"
End Sub

' ---------------------------------------------------------------------------
' Demo rows from the seed folder
' ---------------------------------------------------------------------------

' Parents before children: a child's MsgKey, SubName or GroupName column is
' looked up to its parent's id.
Private Sub SeedDemo(db As DAO.Database, ByVal folder As String)
    Dim t As Variant
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    If Len(Dir(folder & "tblMessages.tsv")) = 0 Then Err.Raise 53, , "No seed files in " & folder
    For Each t In Array("tblRoles", "tblMessages", "tblMessageSections", "tblAttachments", "tblAddressees", "tblAckReceipts", _
            "tblGroupMembers", "tblSubscriptions", "tblPersonSubs", "tblAudit", "tblWatchNotes", "tblOngoingTopics", _
            "tblTurnoverChecklist")
        LoadTsv db, CStr(t), folder & t & ".tsv"
    Next
    SeedDraftingDemo db
End Sub

' Add the rows of one seed file to table tbl (see tools\make_seed.py for the format).
Private Sub LoadTsv(db As DAO.Database, ByVal tbl As String, ByVal path As String)
    Dim lines() As String, cols() As String, vals() As String, i As Long, j As Long, rs As DAO.Recordset
    Dim fld As String, v As Variant
    lines = Split(Replace(ReadUtf8(path), vbCr, ""), vbLf)
    cols = Split(lines(0), vbTab)
    Set rs = db.OpenRecordset(tbl, dbOpenDynaset)
    For i = 1 To UBound(lines)
        If Len(lines(i)) > 0 Then
            vals = Split(lines(i), vbTab)
            If UBound(vals) <> UBound(cols) Then Err.Raise 5, , tbl & " row " & i & " has " & (UBound(vals) + 1) & " values, not " & (UBound(cols) + 1)
            rs.AddNew
            For j = 0 To UBound(cols)
                fld = cols(j)
                If Len(vals(j)) = 0 Then
                    v = Null
                Else
                    v = Unescape(vals(j))
                End If
                Select Case fld
                    Case "MsgKey"
                        If tbl = "tblMessages" Then rs!MsgKey = v Else rs!MessageID = ParentID(db, "tblMessages", "MessageID", "MsgKey", v)
                    Case "SubName"
                        If tbl = "tblSubscriptions" Then rs!SubName = v Else rs!SubID = ParentID(db, "tblSubscriptions", "SubID", "SubName", v)
                    Case "GroupName"
                        rs!GroupID = GroupIdFor(db, CStr(v))
                    Case Else
                        If IsNull(v) Then
                            rs.Fields(fld).Value = Null
                        ElseIf rs.Fields(fld).Type = dbDate Then
                            rs.Fields(fld).Value = IsoDate(CStr(v))
                        ElseIf rs.Fields(fld).Type = dbBoolean Then
                            rs.Fields(fld).Value = (v = "1")
                        Else
                            rs.Fields(fld).Value = v
                        End If
                End Select
            Next
            rs.Update
        End If
    Next
    rs.Close
End Sub

Private Function ParentID(db As DAO.Database, ByVal tbl As String, ByVal idFld As String, ByVal keyFld As String, ByVal key As Variant) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT " & idFld & " FROM " & tbl & " WHERE " & keyFld & "='" & Replace(Nz(key, ""), "'", "''") & "'", dbOpenSnapshot)
    If rs.EOF Then Err.Raise 5, , tbl & " has no " & keyFld & " " & Nz(key, "")
    ParentID = rs.Fields(0).Value
    rs.Close
End Function

' The address group's id, adding the group the first time it is named.
Private Function GroupIdFor(db As DAO.Database, ByVal nm As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT GroupID FROM tblAddressGroups WHERE GroupName='" & Replace(nm, "'", "''") & "'", dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        db.Execute "INSERT INTO tblAddressGroups (GroupName) VALUES ('" & Replace(nm, "'", "''") & "')", dbFailOnError
        Set rs = db.OpenRecordset("SELECT GroupID FROM tblAddressGroups WHERE GroupName='" & Replace(nm, "'", "''") & "'", dbOpenSnapshot)
    End If
    GroupIdFor = rs!GroupID
    rs.Close
End Function

' "yyyy-mm-dd hh:nn:ss" as a date, whatever the computer's date settings.
Private Function IsoDate(ByVal s As String) As Date
    IsoDate = DateSerial(CInt(Mid$(s, 1, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2))) + _
        TimeSerial(CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 15, 2)), CInt(Mid$(s, 18, 2)))
End Function

' \n, \t and \\ back to a line break, a tab and a backslash.
Private Function Unescape(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    If InStr(s, "\") = 0 Then Unescape = s: Exit Function
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "\" And i < Len(s) Then
            Select Case Mid$(s, i + 1, 1)
                Case "n": out = out & vbCrLf
                Case "t": out = out & vbTab
                Case Else: out = out & Mid$(s, i + 1, 1)
            End Select
            i = i + 2
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    Unescape = out
End Function

' A UTF-8 text file's text.
Private Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                                 ' text
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText(-1)
    st.Close
End Function

' ---------------------------------------------------------------------------
' Drafting and release (came over from Message Generator, Stephen 2026-11-05)
' ---------------------------------------------------------------------------
' Outgoing command traffic (the front end's modDrafting, modTemplates and
' modRouting): drafts and their paragraphs, each draft's route (one row per
' step, RouteNo counting the submits), the routing log (tblRoutingEvents,
' which the app only appends to), templates and their field rules, the
' routing workflows and their steps, release authorities, and the one row of
' command transmission settings. Steps and designations name a Personnel
' person_id; the name is kept with a draft's step for the record.

Private Sub CreateDraftingTables(db As DAO.Database)
    ExecDdl db, "CREATE TABLE tblDrafts (DraftID AUTOINCREMENT PRIMARY KEY, Subject TEXT(255), MsgType TEXT(30), MsgClass TEXT(1), " & _
      "Precedence TEXT(12), Classification TEXT(20), Originator TEXT(60), ToLine TEXT(255), InfoLine TEXT(255), TemplateID LONG, " & _
      "DrafterID LONG, DrafterName TEXT(100), Status TEXT(12) NOT NULL, WorkflowID LONG, RouteNo LONG, CreatedOn DATETIME, " & _
      "UpdatedOn DATETIME, Dtg TEXT(20), Msn TEXT(10), LocalSequence TEXT(12), RenderedText MEMO, ReleasedOn DATETIME, " & _
      "ReleasedByID LONG, ReleasedByName TEXT(100), ReleaseDesignation TEXT(100), MessageID LONG)"
    ExecDdl db, "CREATE TABLE tblDraftParagraphs (ParagraphID AUTOINCREMENT PRIMARY KEY, DraftID LONG NOT NULL, SortOrder INTEGER, ParaText MEMO)"
    ExecDdl db, "CREATE TABLE tblDraftRoute (RouteStepID AUTOINCREMENT PRIMARY KEY, DraftID LONG NOT NULL, RouteNo LONG NOT NULL, " & _
      "StepOrder INTEGER NOT NULL, StepTitle TEXT(100), StepAction TEXT(10), PersonID LONG, PersonName TEXT(100), " & _
      "StepStatus TEXT(10) NOT NULL, ActedOn DATETIME, StepComment MEMO)"
    ExecDdl db, "CREATE TABLE tblRoutingEvents (EventID AUTOINCREMENT PRIMARY KEY, DraftID LONG NOT NULL, EventType TEXT(12) NOT NULL, " & _
      "EventAt DATETIME NOT NULL, ActorID LONG, ActorName TEXT(100), StepTitle TEXT(100), EventComment MEMO)"
    ExecDdl db, "CREATE TABLE tblTemplates (TemplateID AUTOINCREMENT PRIMARY KEY, TemplateName TEXT(100) NOT NULL, VersionNo LONG NOT NULL, " & _
      "Description TEXT(255), MsgType TEXT(30), MsgClass TEXT(1), Precedence TEXT(12), Classification TEXT(20), ToLine TEXT(255), " & _
      "InfoLine TEXT(255), Subject TEXT(255), BodyText MEMO, Status TEXT(10) NOT NULL, CreatedOn DATETIME, CreatedBy TEXT(100), " & _
      "UpdatedOn DATETIME, UpdatedBy TEXT(100), LastUsedOn DATETIME)"
    ExecDdl db, "CREATE TABLE tblTemplateFields (FieldID AUTOINCREMENT PRIMARY KEY, TemplateID LONG NOT NULL, FieldName TEXT(64) NOT NULL, " & _
      "SortOrder INTEGER, IsRequired YESNO, MaxLen LONG, DefaultText TEXT(255), Hint TEXT(100))"
    ExecDdl db, "CREATE TABLE tblWorkflows (WorkflowID AUTOINCREMENT PRIMARY KEY, WorkflowName TEXT(100) NOT NULL, Description TEXT(255), " & _
      "IsActive YESNO, MatchTypes TEXT(255), MatchOriginators TEXT(255), UpdatedOn DATETIME, UpdatedBy TEXT(100))"
    ExecDdl db, "CREATE TABLE tblWorkflowSteps (WfStepID AUTOINCREMENT PRIMARY KEY, WorkflowID LONG NOT NULL, StepOrder INTEGER NOT NULL, " & _
      "StepTitle TEXT(100), StepAction TEXT(10) NOT NULL, PersonID LONG)"
    ExecDdl db, "CREATE TABLE tblReleaseAuthorities (AuthorityID AUTOINCREMENT PRIMARY KEY, PersonID LONG NOT NULL, " & _
      "Designation TEXT(100) NOT NULL, IsCommandingOfficer YESNO, Scope TEXT(30), CreatedOn DATETIME, CreatedBy TEXT(100))"
    ExecDdl db, "CREATE TABLE tblCommandSettings (SettingsID LONG PRIMARY KEY, Originator TEXT(60), MsgIdOriginator TEXT(20), " & _
      "TextCase TEXT(5), SeqPrefix TEXT(4), SeqDigits INTEGER, NextSequence LONG, UpdatedOn DATETIME, UpdatedBy TEXT(100))"
End Sub

Private Sub CreateDraftingIndexes(db As DAO.Database)
    ExecDdl db, "CREATE INDEX ixDraftStatus ON tblDrafts (Status, UpdatedOn)"
    ExecDdl db, "CREATE INDEX ixDraftDrafter ON tblDrafts (DrafterID)"
    ExecDdl db, "CREATE INDEX ixParaDraft ON tblDraftParagraphs (DraftID, SortOrder)"
    ExecDdl db, "CREATE INDEX ixRouteDraft ON tblDraftRoute (DraftID, RouteNo, StepOrder)"
    ExecDdl db, "CREATE INDEX ixRoutePerson ON tblDraftRoute (PersonID, StepStatus)"
    ExecDdl db, "CREATE INDEX ixEventDraft ON tblRoutingEvents (DraftID, EventAt)"
    ExecDdl db, "CREATE UNIQUE INDEX uxTemplateVersion ON tblTemplates (TemplateName, VersionNo)"
    ExecDdl db, "CREATE INDEX ixFieldTemplate ON tblTemplateFields (TemplateID, SortOrder)"
    ExecDdl db, "CREATE UNIQUE INDEX uxWorkflowName ON tblWorkflows (WorkflowName)"
    ExecDdl db, "CREATE INDEX ixStepWorkflow ON tblWorkflowSteps (WorkflowID, StepOrder)"
    ExecDdl db, "CREATE INDEX ixAuthorityPerson ON tblReleaseAuthorities (PersonID)"
End Sub

' A draft's paragraphs and route go with it, a template's field rules and a
' workflow's steps with theirs; the routing log never does (the app only
' appends to it, and never deletes a draft).
Private Sub CreateDraftingRelationships(db As DAO.Database)
    AddRelation db, "rDP_D", "tblDrafts", "DraftID", "tblDraftParagraphs", "DraftID", True
    AddRelation db, "rDR_D", "tblDrafts", "DraftID", "tblDraftRoute", "DraftID", True
    AddRelation db, "rRE_D", "tblDrafts", "DraftID", "tblRoutingEvents", "DraftID"
    AddRelation db, "rTF_T", "tblTemplates", "TemplateID", "tblTemplateFields", "TemplateID", True
    AddRelation db, "rWS_W", "tblWorkflows", "WorkflowID", "tblWorkflowSteps", "WorkflowID", True
End Sub

' The command's transmission settings: the FM line, the MSGID originator,
' every message in capitals, and local sequence #0042 next (the web app's).
' The prefix and digits stay # and 4 (the naval format check requires them).
Private Sub SeedDraftingLookups(db As DAO.Database)
    ExecDdl db, "INSERT INTO tblCommandSettings (SettingsID, Originator, MsgIdOriginator, TextCase, SeqPrefix, SeqDigits, NextSequence, " & _
      "UpdatedOn, UpdatedBy) VALUES (1, 'COMSUBLANT NORFOLK VA', 'COMSUBLANT', 'upper', '#', 4, 42, Now(), 'Data file build')"
End Sub

' Demo rows: the web app's four routing workflows (on, their steps held by
' nobody yet: an administrator picks the people from Personnel under Admin:
' Routing) and four fictional templates, active, version 1.
Private Sub SeedDraftingDemo(db As DAO.Database)
    AddWorkflow db, "General command traffic", "COMSUBLANT messages: N3 review, Chief of Staff review, Commander release.", _
        "COMSUBLANT MSG", "", "N3 review|Chief of Staff review|Commander release"
    AddWorkflow db, "Safety traffic", "Safety messages from the staff: N3 review, Chief of Staff release.", _
        "SAFETY MSG", "COMSUBLANT NORFOLK VA", "N3 review|Chief of Staff release"
    AddWorkflow db, "Logistics traffic", "Fleet messages: Supply review, N4 release.", "FLEET MSG", "", "Supply review|N4 release"
    AddWorkflow db, "Default chain", "Any other traffic: Chief of Staff review, Commander release.", "", "", _
        "Chief of Staff review|Commander release"

    AddTemplate db, "General admin notice", "A general administrative notice to subordinate commands.", "GENADMIN", "ROUTINE", _
        "{{ Action addressees }}", "", "{{ Subject }}", _
        "{{ Purpose }}" & vbCrLf & vbCrLf & "{{ Details }}" & vbCrLf & vbCrLf & "POC IS {{ POC }}.", _
        "Action addressees|Commands that act on it, separated by semicolons", _
        "Subject|The SUBJ line, in a few words", "Purpose|What the message is for, in one or two sentences", _
        "Details|What each addressee does, and by when", "POC|Rank, name, code and phone"
    AddTemplate db, "Safety stand-down", "Directs a safety stand-down for one or more units.", "SAFETY MSG", "PRIORITY", _
        "{{ Units }}", "COMSUBLANT NORFOLK VA", "SAFETY STAND-DOWN FOR {{ Units }}", _
        "UNITS LISTED WILL CONDUCT A SAFETY STAND-DOWN NO LATER THAN {{ Complete by }}." & vbCrLf & vbCrLf & _
        "THE STAND-DOWN COVERS {{ Topic }}." & vbCrLf & vbCrLf & "REPORT COMPLETION TO COMSUBLANT N3 BY MESSAGE. POC IS {{ POC }}.", _
        "Units|The units that stand down", "Complete by|A date, such as 15 NOV 26", "Topic|What the stand-down covers", _
        "POC|Rank, name, code and phone"
    AddTemplate db, "Logistics status request", "Asks units for their supply and parts status.", "FLEET MSG", "ROUTINE", _
        "{{ Units }}", "", "LOGISTICS STATUS REQUEST", _
        "REQUEST UNITS LISTED REPORT {{ Items }} STATUS BY {{ Report by }}." & vbCrLf & vbCrLf & _
        "INCLUDE CASREP NUMBERS WHERE THEY APPLY. POC IS {{ POC }}.", _
        "Units|The units that report", "Items|What they report on, such as REPAIR PARTS", "Report by|A date-time group or date", _
        "POC|Rank, name, code and phone"
    AddTemplate db, "Watch bill change", "Tells the staff of a change to the command duty watch bill.", "COMSUBLANT MSG", "ROUTINE", _
        "{{ Action addressees }}", "", "COMMAND DUTY WATCH BILL CHANGE", _
        "EFFECTIVE {{ Effective }}, {{ Change }}." & vbCrLf & vbCrLf & "POC IS {{ POC }}.", _
        "Action addressees|Commands or codes that act on it", "Effective|The date the change starts", _
        "Change|Who stands which watch, in a sentence", "POC|Rank, name, code and phone"
End Sub

' A workflow, on, with the steps named in steps ("A|B|C"), the last one the
' release step, nobody holding them.
Private Sub AddWorkflow(db As DAO.Database, ByVal nm As String, ByVal desc As String, ByVal types As String, ByVal origs As String, _
        ByVal steps As String)
    Dim id As Long, p() As String, i As Long
    ExecDdl db, "INSERT INTO tblWorkflows (WorkflowName, Description, IsActive, MatchTypes, MatchOriginators, UpdatedOn, UpdatedBy) VALUES ('" & _
        nm & "','" & desc & "',True,'" & types & "','" & origs & "',Now(),'Data file build')"
    id = db.OpenRecordset("SELECT @@IDENTITY")(0)
    p = Split(steps, "|")
    For i = 0 To UBound(p)
        ExecDdl db, "INSERT INTO tblWorkflowSteps (WorkflowID, StepOrder, StepTitle, StepAction) VALUES (" & id & "," & (i + 1) & ",'" & _
            p(i) & "','" & IIf(i = UBound(p), "release", "review") & "')"
    Next
End Sub

' A template, active, version 1, UNCLASSIFIED, class A, with its fields in
' order, each "Name|hint": required, no length limit.
Private Sub AddTemplate(db As DAO.Database, ByVal nm As String, ByVal desc As String, ByVal msgType As String, ByVal prec As String, _
        ByVal toLine As String, ByVal infoLine As String, ByVal subj As String, ByVal body As String, ParamArray fields() As Variant)
    Dim id As Long, i As Long, p() As String
    ExecDdl db, "INSERT INTO tblTemplates (TemplateName, VersionNo, Description, MsgType, MsgClass, Precedence, Classification, ToLine, " & _
        "InfoLine, Subject, BodyText, Status, CreatedOn, CreatedBy, UpdatedOn, UpdatedBy) VALUES ('" & nm & "',1,'" & desc & "','" & _
        msgType & "','A','" & prec & "','UNCLASSIFIED','" & toLine & "','" & infoLine & "','" & subj & "','" & body & _
        "','Active',Now(),'Data file build',Now(),'Data file build')"
    id = db.OpenRecordset("SELECT @@IDENTITY")(0)
    For i = 0 To UBound(fields)
        p = Split(fields(i), "|")
        ExecDdl db, "INSERT INTO tblTemplateFields (TemplateID, FieldName, SortOrder, IsRequired, MaxLen, DefaultText, Hint) VALUES (" & _
            id & ",'" & p(0) & "'," & (i + 1) & ",True,0,'','" & p(1) & "')"
    Next
End Sub

' ---------------------------------------------------------------------------
' Mail ingestion (a prototype, Stephen 2026-11-05)
' ---------------------------------------------------------------------------
' Received traffic read from a dedicated mailbox (Outlook) or a drop folder
' of .txt files by the front end's modAnchorIngest. tblIngestLog has one row
' per attempt, whatever its result (stored, repeat, conflict, rejected), and
' the app only appends to it; MessageID names the message an attempt stored
' or matched. The raw text of a stored message is tblMessages.RawText (with
' its SHA-256 in Sha256), so there is no separate raw-text table; refused
' text is kept by the feeders in a Rejected folder, never in the data file.
' The settings are tblSettings rows (blank ones Null, so the app's defaults
' apply); IngestLastRun, the last run's counts, is added by the first run.

Private Sub CreateIngestTables(db As DAO.Database)
    ExecDdl db, "CREATE TABLE tblIngestLog (IngestID AUTOINCREMENT PRIMARY KEY, LoggedAt DATETIME NOT NULL, Workstation TEXT(64), " & _
      "UserName TEXT(100), Feeder TEXT(20) NOT NULL, EmailMessageId TEXT(255), Reference TEXT(60), Sha256 TEXT(64), " & _
      "Result TEXT(10) NOT NULL, Reason TEXT(255), FileOrItem TEXT(255), MessageID LONG)"
End Sub

Private Sub CreateIngestIndexes(db As DAO.Database)
    ExecDdl db, "CREATE INDEX ixIngestAt ON tblIngestLog (LoggedAt)"
    ExecDdl db, "CREATE INDEX ixIngestRef ON tblIngestLog (Reference)"
End Sub

' The drop folder blank (the "ingest" folder beside the data file), no
' mailbox yet, its Inbox, class A, up to SECRET (demonstration markings), no
' automatic checks.
Private Sub SeedIngestSettings(db As DAO.Database)
    ExecDdl db, "INSERT INTO tblSettings VALUES ('IngestDropFolder',Null,'Mail ingest: the drop folder (blank: ingest beside the data file)')"
    ExecDdl db, "INSERT INTO tblSettings VALUES ('IngestMailbox',Null,'Mail ingest: the mailbox Outlook reads')"
    ExecDdl db, "INSERT INTO tblSettings VALUES ('IngestFolder','Inbox','Mail ingest: the folder in that mailbox')"
    ExecDdl db, "INSERT INTO tblSettings VALUES ('IngestDefaultClass','A','Mail ingest: the class received messages take')"
    ExecDdl db, "INSERT INTO tblSettings VALUES ('IngestMaxClassification','SECRET','Mail ingest: the highest marking taken')"
    ExecDdl db, "INSERT INTO tblSettings VALUES ('IngestTimerMinutes','0','Mail ingest: minutes between checks while Admin is open (0 off)')"
End Sub
