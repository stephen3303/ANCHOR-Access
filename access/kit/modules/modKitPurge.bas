Option Compare Database
Option Explicit

' Access UI kit - an administrator's "delete all data" (purge) of the shared
' data file, with the safeguards a step that can't be undone needs.
'
' The app's Admin page, for an administrator only, calls:
'   If KitPurgeData("command", PurgeTables(), "Every inspection, grade ...", _
'       "Settings, user accounts ...") Then  ' reseed the essentials, sign out, restart
' PurgeTables() is the app's own list of tables to empty, children before
' parents. KitPurgeData:
'   1. refuses when another computer has the data file open
'      (KitDataFileUsers). A single-file copy, whose tables are its own, is
'      its own data file;
'   2. shows dlgConfirm (KitConfirmTyped): a red warning that says what will be
'      deleted, what is kept and where the backup goes, and a box where the
'      Delete button stays disabled until PURGE_WORD is typed exactly, as
'      GitHub asks for a repository's name before deleting it;
'   3. copies the data file to a Backups folder beside it (KitBackupDataFile),
'      checks the copy opens with every table, and stops if it doesn't (the
'      user may pick another folder once);
'   4. empties the tables in one transaction (KitPurgeTables): every table or
'      none;
'   5. writes a warning event to the Windows Event Log (EVENT_ID_PURGE), which
'      the purge can't delete: who, which file, how many rows, the backup.
'   6. tells the user what was deleted, where the backup is, and that the app
'      closes now.
' It returns True only when the tables were emptied. The app then puts back
' any rows it needs to open (its own first-run settings) and quits
' (DoCmd.Quit), so the next start opens cleanly: signed in again, or at the
' Personnel setup screen after a Personnel wipe.
'
' The shared Personnel file is a separate file that every app shares, and its
' tables are never in an app's list. The dialog offers deleting its people
' too, as a check box that starts unticked. Ticked, the purge also refuses
' while any other computer has the Personnel file open (any app, anywhere),
' saves a checked backup of it (with its Personnel.key) in a Backups folder
' beside it, and empties kit_person_app and kit_person in place, so every
' app's links and the site key still work. The admin is signed out
' (KitPurgedPersonnel is True), and the next start of any app shows the
' Personnel setup screen to create the first admin (PERSONNEL_NO_ADMIN).
'
' The app provides APP_NAME and builds dlgConfirm (BuildConfirmDialog, with
' kit\forms\dlgConfirm.vba copied to src\forms).

Public Const PURGE_WORD As String = "DELETE"
Public Const EVENT_ID_PURGE As Long = 1002
Public Const BACKUP_FOLDER As String = "Backups"

' The question dlgConfirm is showing (KitConfirmTyped).
Private m_cfTitle As String
Private m_cfHeading As String
Private m_cfDetail As String
Private m_cfWord As String
Private m_cfAction As String
Private m_cfOption As String          ' the check box's text, "" for none
Private m_cfOptionChosen As Boolean   ' it was ticked when the action was clicked
Private m_purgedPersonnel As Boolean  ' the last KitPurgeData emptied the Personnel file too

' ---------------------------------------------------------------------------
' Type-to-confirm dialog
' ---------------------------------------------------------------------------

' Ask before something that can't be undone: dlgConfirm shows heading in red,
' detail (what happens, what is kept), and a box where the user types word;
' the action button (caption action, in the destructive color) stays disabled
' until the box holds word exactly (case counts). optionText, when given,
' shows a check box (unticked) with that text; KitConfirmOptionChosen says
' whether it was ticked. True only when that button was clicked. Cancel,
' Esc, the window's X, or a dialog that can't open (not built yet) all
' return False: a step like this is never confirmed by a fallback.
Public Function KitConfirmTyped(ByVal heading As String, ByVal detail As String, _
        Optional ByVal word As String = PURGE_WORD, Optional ByVal action As String = "Delete", _
        Optional ByVal title As String = "", Optional ByVal optionText As String = "") As Boolean
    On Error GoTo Failed
    SetKitConfirm heading, detail, word, action, title, optionText
    CloseSplash
    KitConfirmTyped = RunDialog("dlgConfirm")
    ClearConfirm
    Exit Function
Failed:
    LogError "modKitPurge.KitConfirmTyped", heading
    ClearConfirm
End Function

' Set what dlgConfirm shows next (KitConfirmTyped; the kit tests open the
' dialog themselves after it).
Public Sub SetKitConfirm(ByVal heading As String, ByVal detail As String, ByVal word As String, _
        ByVal action As String, Optional ByVal title As String = "", Optional ByVal optionText As String = "")
    m_cfTitle = IIf(Len(title) > 0, title, APP_NAME)
    m_cfOption = optionText
    m_cfOptionChosen = False
    m_cfHeading = heading
    m_cfDetail = detail
    m_cfWord = word
    m_cfAction = action
End Sub

' What the open dlgConfirm shows (its Form_Load calls these).
Public Function KitConfirmTitle() As String
    KitConfirmTitle = m_cfTitle
End Function

Public Function KitConfirmHeading() As String
    KitConfirmHeading = m_cfHeading
End Function

Public Function KitConfirmDetail() As String
    KitConfirmDetail = m_cfDetail
End Function

Public Function KitConfirmWord() As String
    KitConfirmWord = m_cfWord
End Function

Public Function KitConfirmAction() As String
    KitConfirmAction = m_cfAction
End Function

Public Function KitConfirmOption() As String
    KitConfirmOption = m_cfOption
End Function

' Whether the last confirmed question's check box was ticked (dlgConfirm sets it).
Public Function KitConfirmOptionChosen() As Boolean
    KitConfirmOptionChosen = m_cfOptionChosen
End Function

Public Sub SetKitConfirmOptionChosen(ByVal chosen As Boolean)
    m_cfOptionChosen = chosen
End Sub

' True when typed is the word exactly: same letters, same case, no spaces
' around it.
Public Function KitConfirmMatches(ByVal typed As String) As Boolean
    If Len(m_cfWord) = 0 Then Exit Function
    KitConfirmMatches = (StrComp(typed, m_cfWord, vbBinaryCompare) = 0)
End Function

Private Sub ClearConfirm()
    m_cfTitle = "": m_cfHeading = "": m_cfDetail = "": m_cfWord = "": m_cfAction = "": m_cfOption = ""
End Sub

' ---------------------------------------------------------------------------
' Purge
' ---------------------------------------------------------------------------

' The whole purge (see the top of the module). probeTable is one of the app's
' linked tables; tables (an array) are emptied in order, children first;
' deleted and kept describe them for the warning, in the user's words.
' offerPersonnel:=False leaves out the Personnel check box (it is also left
' out when no Personnel file is linked). Every refusal and failure is
' explained to the user. Returns True when the tables were emptied;
' KitPurgedPersonnel then says whether the Personnel file was emptied too.
Public Function KitPurgeData(ByVal probeTable As String, ByVal tables As Variant, ByVal deleted As String, _
        ByVal kept As String, Optional ByVal offerPersonnel As Boolean = True) As Boolean
    Dim bePath As String, others As String, folder As String, backup As String, rows As Long
    Dim busy As Boolean, stage As String, pePath As String, peBackup As String, people As Boolean, peRows As Long
    Dim optionText As String
    On Error GoTo Fail
    stage = "check"
    m_purgedPersonnel = False
    bePath = KitBackendPath(probeTable)
    If Len(bePath) = 0 Then
        ' A single-file copy keeps its tables itself: it is its own data file.
        If Not IsLocalTable(probeTable) Then
            KitMsgBox "This copy of " & APP_NAME & " has no " & probeTable & " table, so there is nothing to purge.", _
                vbInformation
            Exit Function
        End If
        bePath = CurrentDb.Name
    End If
    others = KitDataFileUsers(bePath)
    If Len(others) > 0 Then
        KitMsgBox "The data can't be deleted while anyone else has " & APP_NAME & " open." & vbCrLf & vbCrLf & _
            "Open on: " & others & vbCrLf & vbCrLf & "Ask them to close it, then try again.", vbExclamation
        Exit Function
    End If
    folder = DefaultBackupFolder(bePath)
    If offerPersonnel Then pePath = PersonnelPath()
    If Len(pePath) > 0 Then
        optionText = "Also delete everyone in the shared Personnel directory. This signs everyone out of every app " & _
            "that uses it; the next person to open one creates the first admin."
    End If
    If Not KitConfirmTyped("Delete all " & APP_NAME & " data", _
            "This deletes, for every user:" & vbCrLf & deleted & vbCrLf & vbCrLf & _
            "It keeps:" & vbCrLf & kept & vbCrLf & vbCrLf & _
            "A backup copy of the data file is saved first, in:" & vbCrLf & folder & vbCrLf & _
            IIf(Len(optionText) > 0, "(and of the Personnel file beside it, when the box below is ticked)" & vbCrLf, "") & _
            vbCrLf & _
            "This can't be undone from inside " & APP_NAME & ". Only the backup copy can bring the data back.", _
            PURGE_WORD, "Delete all data", "", optionText) Then Exit Function
    people = (Len(pePath) > 0 And KitConfirmOptionChosen())
    If people Then
        others = KitDataFileUsers(pePath)
        If Len(others) > 0 Then
            KitMsgBox "Nothing was deleted: the Personnel directory is open in an app on " & others & "." & vbCrLf & _
                vbCrLf & "Ask them to close every app that uses it, then try again.", vbExclamation
            Exit Function
        End If
        KitBusy "Saving a backup copy of the Personnel file..."
        busy = True
        peBackup = TryBackup(pePath, DefaultBackupFolder(pePath))
        If Len(peBackup) > 0 Then peBackup = WithKeyCopy(pePath, peBackup)
        KitBusyEnd
        busy = False
        If Len(peBackup) = 0 Then
            KitMsgBox "Nothing was deleted: the backup copy of the Personnel file couldn't be saved in the " & _
                BACKUP_FOLDER & " folder beside it.", vbExclamation
            Exit Function
        End If
    End If

    stage = "backup"
    KitBusy "Saving a backup copy of the data file..."
    busy = True
    backup = TryBackup(bePath, folder)
    KitBusyEnd
    busy = False
    If Len(backup) = 0 Then
        KitMsgBox "The backup copy couldn't be saved in the Backups folder beside the data file." & vbCrLf & vbCrLf & _
            "Pick another folder for it in the next window, or Cancel to stop. Nothing has been deleted.", vbExclamation
        folder = KitPickFolder("Pick a folder for the backup copy", "purge-backup")
        If Len(folder) = 0 Then
            KitMsgBox "Nothing was deleted: the backup copy couldn't be saved.", vbExclamation
            Exit Function
        End If
        KitBusy "Saving a backup copy of the data file..."
        busy = True
        backup = TryBackup(bePath, folder)
        KitBusyEnd
        busy = False
        If Len(backup) = 0 Then
            KitMsgBox "Nothing was deleted: the backup copy couldn't be saved in that folder either.", vbExclamation
            Exit Function
        End If
    End If

    ' Someone may have opened the app while the user was typing.
    others = KitDataFileUsers(bePath)
    If Len(others) > 0 Then
        KitMsgBox "Nothing was deleted: " & APP_NAME & " was opened on " & others & " meanwhile." & vbCrLf & vbCrLf & _
            "Ask them to close it, then try again.", vbExclamation
        Exit Function
    End If
    If people Then others = KitDataFileUsers(pePath)
    If Len(others) > 0 Then
        KitMsgBox "Nothing was deleted: an app using the Personnel directory was opened on " & others & " meanwhile." & _
            vbCrLf & vbCrLf & "Ask them to close it, then try again.", vbExclamation
        Exit Function
    End If

    stage = "delete"
    KitBusy "Deleting the data..."
    busy = True
    rows = KitPurgeTables(bePath, tables, ConnectPassword(probeTable))
    KitBusyEnd
    busy = False
    ' Committed: from here on an error must not say nothing was deleted.
    stage = "deleted"
    KitPurgeData = True
    KitEventLog APP_NAME & ": " & KitUser() & " (Windows " & WinAccount() & " on " & WinComputer() & _
        ") deleted all data in " & bePath & ": " & rows & " rows from " & (UBound(tables) - LBound(tables) + 1) & _
        " tables." & vbCrLf & "Backup copy: " & backup, EVENT_TYPE_WARNING, EVENT_ID_PURGE
    If people Then
        stage = "personnel"
        KitBusy "Deleting the people in the Personnel directory..."
        busy = True
        peRows = KitPurgeTables(pePath, Array(PERSON_APP_TABLE, PERSON_TABLE), PersonnelKey())
        KitBusyEnd
        busy = False
        m_purgedPersonnel = True
        stage = "deleted"
        PersonnelSignOut
        KitEventLog APP_NAME & ": " & WinAccount() & " on " & WinComputer() & " deleted everyone in the Personnel " & _
            "directory " & pePath & ": " & peRows & " rows." & vbCrLf & "Backup copy: " & peBackup, _
            EVENT_TYPE_WARNING, EVENT_ID_PURGE
    End If
    KitMsgBox "All " & APP_NAME & " data was deleted" & IIf(people, ", and everyone in the Personnel directory", "") & _
        "." & vbCrLf & vbCrLf & "The backup copy is in the " & BACKUP_FOLDER & " folder beside the data file." & _
        vbCrLf & vbCrLf & APP_NAME & " closes now. Open it again to carry on" & _
        IIf(people, ": the first person to open an app creates the first admin.", "."), vbInformation
    Exit Function
Fail:
    If busy Then KitBusyEnd
    Select Case stage
        Case "personnel"
            HandleError "modKitPurge.KitPurgeData", pePath, "The " & APP_NAME & " data was deleted, but the Personnel " & _
                "directory couldn't be emptied, so nobody was removed from it."
        Case "deleted"
            HandleError "modKitPurge.KitPurgeData", bePath, "The " & APP_NAME & " data was deleted" & _
                IIf(m_purgedPersonnel, ", and everyone in the Personnel directory", _
                IIf(people, ", but the Personnel directory wasn't emptied", "")) & ". The backup copy is in " & backup & "."
        Case "delete"
            HandleError "modKitPurge.KitPurgeData", bePath, "Nothing was deleted: the data couldn't be deleted, " & _
                "so every table was left as it was. The backup copy is in " & backup & "."
        Case Else
            HandleError "modKitPurge.KitPurgeData", bePath, "Nothing was deleted."
    End Select
End Function

' True when the last KitPurgeData emptied the Personnel file too: the app
' then restarts into the Personnel setup screen, not its sign-in.
Public Function KitPurgedPersonnel() As Boolean
    KitPurgedPersonnel = m_purgedPersonnel
End Function

' The computers other than this one that have the data file at bePath open, as
' "PC1, PC2", or "" when none. Read from the file's user roster (the users
' its .laccdb lock file lists), so a computer that left the file open after a
' crash is listed until the lock file clears. Two users on one computer (a
' terminal server) show as that computer, which is this one's name for a user
' on the same server, so they aren't seen: purge from a PC, not a server.
' Raises when the roster can't be read.
Public Function KitDataFileUsers(ByVal bePath As String) As String
    Const JET_SCHEMA_USERROSTER As String = "{947bb102-5d43-11d1-bdbf-00c04fb92675}"
    Const adSchemaProviderSpecific As Long = -1
    Dim cn As Object, rs As Object, pc As String, thisPc As String, seen As Object, pwd As String
    On Error GoTo Fail
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    thisPc = WinComputer()
    pwd = PasswordOfFile(bePath)
    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & bePath & _
        IIf(Len(pwd) > 0, ";Jet OLEDB:Database Password=" & pwd, "")
    Set rs = cn.OpenSchema(adSchemaProviderSpecific, , JET_SCHEMA_USERROSTER)
    Do Until rs.EOF
        pc = Trim$(Replace(Nz(rs.Fields(0).Value, ""), Chr$(0), ""))
        If Len(pc) > 0 And StrComp(pc, thisPc, vbTextCompare) <> 0 And Not seen.Exists(pc) Then
            seen.Add pc, True
            KitDataFileUsers = KitDataFileUsers & IIf(Len(KitDataFileUsers) > 0, ", ", "") & pc
        End If
        rs.MoveNext
    Loop
    rs.Close
    cn.Close
    Exit Function
Fail:
    RaiseAgain "modKitPurge.KitDataFileUsers", bePath
End Function

' Copy the data file at bePath into folder (made if needed) as
' "<name>_backup_<yyyymmdd_hhnnss>.accdb", then check the copy opens and has
' every table the original has. Returns the copy's path; raises (deleting a
' bad copy) when the copy fails or doesn't check out.
Public Function KitBackupDataFile(ByVal bePath As String, ByVal folder As String) As String
    Dim fso As Object, dest As String, src As DAO.Database, cp As DAO.Database, pwd As String
    Dim want As Long, got As Long, copied As Boolean, n As Long, d As String, s As String
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureFolder folder
    dest = folder & "\" & fso.GetBaseName(bePath) & "_backup_" & Format$(Now(), "yyyymmdd_hhnnss") & ".accdb"
    If FileExists(dest) Then Err.Raise 58, "modKitPurge", "A backup with that name already exists: " & dest
    fso.CopyFile bePath, dest, False
    copied = True
    pwd = PasswordOfFile(bePath)
    Set src = DBEngine.OpenDatabase(bePath, False, True, IIf(Len(pwd) > 0, ";PWD=" & pwd, ""))
    want = UserTableCount(src)
    src.Close: Set src = Nothing
    Set cp = DBEngine.OpenDatabase(dest, False, True, IIf(Len(pwd) > 0, ";PWD=" & pwd, ""))
    got = UserTableCount(cp)
    cp.Close: Set cp = Nothing
    If got <> want Or got = 0 Then Err.Raise 5, "modKitPurge", "The backup copy has " & got & " of " & want & " tables."
    KitBackupDataFile = dest
    Exit Function
Fail:
    n = Err.Number: d = Err.Description: s = Err.Source
    On Error Resume Next
    If Not src Is Nothing Then src.Close
    If Not cp Is Nothing Then cp.Close
    If copied Then Kill dest                    ' only the copy this call made
    On Error GoTo 0
    Err.Number = n: Err.Description = d: Err.Source = s
    RaiseAgain "modKitPurge.KitBackupDataFile", bePath & " -> " & folder
End Function

' Delete every row of tables (an array, emptied in that order: children
' before parents) in the data file at bePath, in one transaction: all of them
' or, on any error, none. pwd is the file's password, if any. Returns how
' many rows were deleted.
Public Function KitPurgeTables(ByVal bePath As String, ByVal tables As Variant, Optional ByVal pwd As String = "") As Long
    Dim ws As DAO.Workspace, db As DAO.Database, i As Long, inTrans As Boolean, n As Long, tbl As String
    Dim en As Long, ed As String, es As String
    On Error GoTo Fail
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(bePath, False, False, IIf(Len(pwd) > 0, ";PWD=" & pwd, ""))
    ws.BeginTrans
    inTrans = True
    For i = LBound(tables) To UBound(tables)
        tbl = tables(i)
        db.Execute "DELETE FROM [" & tbl & "]", dbFailOnError
        n = n + db.RecordsAffected
    Next
    ws.CommitTrans dbForceOSFlush
    inTrans = False
    db.Close
    KitPurgeTables = n
    Exit Function
Fail:
    en = Err.Number: ed = Err.Description: es = Err.Source
    On Error Resume Next
    If inTrans Then ws.Rollback
    If Not db Is Nothing Then db.Close
    On Error GoTo 0
    Err.Number = en: Err.Description = ed: Err.Source = es
    RaiseAgain "modKitPurge.KitPurgeTables", bePath & IIf(Len(tbl) > 0, " (" & tbl & ")", "")
End Function

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

' True when the current database has probeTable as a table of its own (not linked).
Private Function IsLocalTable(ByVal probeTable As String) As Boolean
    Dim cn As String
    On Error GoTo Missing
    cn = CurrentDb.TableDefs(probeTable).Connect
    IsLocalTable = (Len(cn) = 0)
Missing:
End Function

' The Backups folder beside the data file.
Private Function DefaultBackupFolder(ByVal bePath As String) As String
    DefaultBackupFolder = Left$(bePath, InStrRev(bePath, "\")) & BACKUP_FOLDER
End Function

' KitBackupDataFile into folder, or "" (logged) when it failed.
Private Function TryBackup(ByVal bePath As String, ByVal folder As String) As String
    On Error GoTo Failed
    TryBackup = KitBackupDataFile(bePath, folder)
    Exit Function
Failed:
    LogError "modKitPurge.TryBackup", folder
End Function

' Copy the Personnel.key beside pePath next to its backup (as
' "<backup name>.key"), since the backup opens only with that key. Returns
' backup, or "" (logged, the backup deleted) when the key couldn't be copied.
Private Function WithKeyCopy(ByVal pePath As String, ByVal backup As String) As String
    Dim kf As String
    On Error GoTo Failed
    kf = PersonnelKeyPath(pePath)
    If FileExists(kf) Then FileCopy kf, Left$(backup, Len(backup) - Len(".accdb")) & ".key"
    WithKeyCopy = backup
    Exit Function
Failed:
    LogError "modKitPurge.WithKeyCopy", kf
    On Error Resume Next
    Kill backup
End Function

' The password in a linked table's Connect ("" when it has none).
Private Function ConnectPassword(ByVal probeTable As String) As String
    Dim parts() As String, i As Long
    On Error Resume Next
    parts = Split(CurrentDb.TableDefs(probeTable).Connect, ";")
    On Error GoTo 0
    For i = 0 To UBound(parts)
        If UCase$(Left$(parts(i), 4)) = "PWD=" Then ConnectPassword = Mid$(parts(i), 5): Exit Function
    Next
End Function

' The password of the links to the data file at bePath ("" when none).
Private Function PasswordOfFile(ByVal bePath As String) As String
    Dim t As Collection
    Set t = KitLinkedTables(bePath)
    If t.Count > 0 Then PasswordOfFile = ConnectPassword(t(1))
End Function

' Tables in db other than Access's own (MSys*) and temporary ones (~*).
Private Function UserTableCount(db As DAO.Database) As Long
    Dim td As DAO.TableDef
    For Each td In db.TableDefs
        If Left$(td.Name, 4) <> "MSys" And Left$(td.Name, 1) <> "~" Then UserTableCount = UserTableCount + 1
    Next
End Function
