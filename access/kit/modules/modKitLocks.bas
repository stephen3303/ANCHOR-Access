Option Compare Database
Option Explicit

' Access UI kit - several users on one back end: record locks and write
' conflicts handled the same way in every app, in the kit's look.
'
' Two kinds of trouble come up once a back end is shared:
'   - Locks. Another user (or another of this user's windows) is writing the
'     same record or table this instant: errors 3008, 3009, 3186-3188, 3197,
'     3202, 3211, 3212, 3218, 3260-3262. They clear in a moment, so the kit
'     retries a few times, waiting a little longer each time, before it
'     gives up with a plain message.
'   - Write conflicts. Another user saved the record after this user started
'     editing it (Access's "Write Conflict" window, form error 7787, or 7878
'     "The data has been changed"). Retrying can't help; the user decides.
'
' For code:
'   n = KitExecute("UPDATE ...")       ' an action query, retried on locks
'   KitSaveRecord Me                   ' save a bound form's record, retried
' and, around DAO edits of your own:
'       attempt = 0
'   TryAgain:
'       On Error GoTo Fail
'       rs.Edit: rs!status = "Closed": rs.Update
'       Exit Sub
'   Fail:
'       If LockWait(Err.Number, attempt) Then Resume TryAgain
'       HandleError "pgItems.CloseItem"
'
' For bound forms, add these to the form's code-behind:
'   Private Sub Form_Current()
'       KitNoteRecord Me
'   End Sub
'   Private Sub Form_AfterUpdate()
'       KitNoteRecord Me               ' the saved record is the new "as loaded"
'   End Sub
'   Private Sub Form_BeforeUpdate(Cancel As Integer)
'       KitNoteEdits Me                ' (with any checks of your own)
'   End Sub
'   Private Sub Form_Error(DataErr As Integer, Response As Integer)
'       KitFormError Me, DataErr, Response
'   End Sub
' and save with KitSaveRecord (the Save button, Ctrl+S). On a write conflict
' the user is asked in a KitMsgBox, naming the fields they changed: Yes puts
' their changes back over the other user's saved record (the other user's
' changes to other fields stay), for them to save again; No drops theirs.
' A save Access starts itself (moving to another record, closing) keeps
' Access's own Write Conflict window: the kit can't safely ask from inside
' it. A lock gets "wait a moment, then save again", with the changes kept.
' Other errors are left to Access.

Public Const LOCK_TRIES As Long = 5            ' retries after the first try
Public Const LOCK_WAIT_MS As Long = 120        ' first wait; each retry waits longer
Public Const ERR_WRITE_CONFLICT As Long = 7787 ' Access's Write Conflict window
Public Const ERR_DATA_CHANGED As Long = 7878   ' "The data has been changed"

Private Const LOCK_MSG As String = "Someone else is saving this record right now. Wait a moment, then save again."
Private Const LOCK_RAISE As String = "Another user is saving the same records right now. Wait a moment, then try again."
Private Const LOCK_BUSY As String = "Waiting for another user to finish saving..."

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private m_testAnswer As Long    ' tests only: the conflict answer, 0 to ask
Private m_inSave As Boolean     ' KitSaveRecord is running (its retries wait with DoEvents)
Private m_saving As Boolean     ' KitSaveRecord is saving: KitFormError leaves locks and conflicts to it
Private m_formLock As Long      ' the lock error KitFormError passed on to KitSaveRecord
Private m_conflictForm As String ' the form whose save hit a write conflict, for SaveTry to resolve
' The record as loaded, noted by KitNoteRecord (Form_Current and
' Form_AfterUpdate): by the time of BeforeUpdate or Form_Error, OldValue
' already holds the other user's saved values, so it can't tell the user's
' edits from theirs. One note per open form, keyed by form name; each is
' Array(bookmark, values keyed by control name).
Private m_recs As Collection
' The user's edits, noted by KitNoteEdits before each save.
Private m_editForm As String
Private m_names As Collection   ' changed controls
Private m_fields As Collection  ' what the user calls them, for the message
Private m_vals As Collection    ' the user's values

' True for the errors a moment's wait can clear (see the top of the module).
Public Function IsLockError(ByVal errNum As Long) As Boolean
    Select Case errNum
        Case 3008, 3009, 3186, 3187, 3188, 3197, 3202, 3211, 3212, 3218, 3260, 3261, 3262
            IsLockError = True
    End Select
End Function

' In an error handler: when errNum is a lock error and fewer than LOCK_TRIES
' retries have been made, wait (longer each time, with a random part so two
' users don't retry in step), count the retry in attempt and return True, so
' the handler resumes and tries again. Otherwise return False.
Public Function LockWait(ByVal errNum As Long, ByRef attempt As Long) As Boolean
    If Not IsLockError(errNum) Then Exit Function
    If attempt >= LOCK_TRIES Then Exit Function
    attempt = attempt + 1
    Randomize
    Pause LOCK_WAIT_MS * attempt + CLng(Rnd() * LOCK_WAIT_MS)
    LockWait = True
End Function

' Run an action query on the current database with dbFailOnError, retrying
' while its records are locked. Returns the records it changed. When the
' locks don't clear, raises the lock error with a plain description for the
' caller's HandleError to show; any other error is raised unchanged.
Public Function KitExecute(ByVal sql As String) As Long
    Dim db As DAO.Database, attempt As Long, num As Long, d As String, src As String, busy As Boolean
    Set db = CurrentDb
TryAgain:
    On Error GoTo Fail
    db.Execute sql, dbFailOnError
    KitExecute = db.RecordsAffected
    If busy Then KitBusyEnd
    Exit Function
Fail:
    ' KitBusy and KitBusyEnd reset Err, so keep the error to pass it on.
    num = Err.Number: d = Err.Description: src = Err.Source
    ' The first retry: say why the app is waiting until the retries end.
    If attempt = 0 And IsLockError(num) Then KitBusy LOCK_BUSY: busy = True
    If LockWait(num, attempt) Then Resume TryAgain
    If busy Then KitBusyEnd
    Err.Number = num: Err.Description = d: Err.Source = src
    If IsLockError(num) Then
        LogError "modKitLocks.KitExecute", "locked after " & LOCK_TRIES & " retries: " & Left$(sql, 200)
        On Error GoTo 0
        Err.Raise num, "modKitLocks.KitExecute", LOCK_RAISE
    End If
    RaiseAgain "modKitLocks.KitExecute", Left$(sql, 200)
End Function

' Save frm's changed record, retrying while it is locked. Returns True when
' the record is saved or had no changes. Returns False when the form refused
' (its BeforeUpdate cancelled), the user chose in KitFormError, or the lock
' didn't clear (the user is told). Other errors are raised.
Public Function KitSaveRecord(frm As Form) As Boolean
    If m_inSave Then Exit Function              ' a click during a retry's wait: the first save carries on
    m_inSave = True
    On Error GoTo Fail
    KitSaveRecord = SaveTry(frm)
    m_inSave = False
    Exit Function
Fail:
    m_inSave = False
    RaiseAgain "modKitLocks.KitSaveRecord", frm.Name
End Function

' A bound form's Form_Error (see the top of the module): handles write
' conflicts and locks in the kit's look and sets Response so Access shows
' nothing; leaves every other error to Access.
Public Sub KitFormError(frm As Form, ByVal DataErr As Integer, Response As Integer)
    On Error GoTo Fail
    Select Case DataErr
        Case ERR_WRITE_CONFLICT, ERR_DATA_CHANGED
            ' Only a KitSaveRecord save, with the edits noted: SaveTry asks once
            ' the failed save is over. Any other save keeps Access's window.
            If m_saving And m_editForm = frm.Name Then
                Response = acDataErrContinue
                m_conflictForm = frm.Name
            End If
        Case Else
            If IsLockError(DataErr) Then
                Response = acDataErrContinue
                If m_saving Then
                    m_formLock = DataErr        ' KitSaveRecord retries, and tells the user if it can't
                Else
                    KitMsgBox LOCK_MSG, vbExclamation
                End If
            End If
    End Select
    Exit Sub
Fail:
    LogError "modKitLocks.KitFormError", frm.Name & " error " & DataErr
End Sub

' A bound form's Current (see the top of the module): note the record's
' bound values as loaded, for KitNoteEdits to compare against.
Public Sub KitNoteRecord(frm As Form)
    Dim c As Control, vals As Collection, bm As Variant
    On Error GoTo Fail
    If m_recs Is Nothing Then Set m_recs = New Collection
    ForgetRecord frm.Name
    If frm.NewRecord Then Exit Sub      ' nothing loaded: every value is the user's
    bm = frm.Bookmark
    Set vals = New Collection
    For Each c In frm.Controls
        If IsBoundEditor(c) Then vals.Add c.Value, c.Name
    Next
    m_recs.Add Array(bm, vals), frm.Name
    Exit Sub
Fail:
    ForgetRecord frm.Name               ' no bookmark: a conflict keeps Access's window
End Sub

' Drop the note for form nm, if there is one.
Private Sub ForgetRecord(ByVal nm As String)
    On Error Resume Next                ' no note for nm
    m_recs.Remove nm
End Sub

' A bound form's BeforeUpdate (see the top of the module): note which
' bound controls the user changed (against the record as KitNoteRecord
' noted it) and their values, so a write conflict can name them and, on
' Yes, put back only those.
Public Sub KitNoteEdits(frm As Form)
    Dim c As Control, src As String, loaded As Boolean, was As Variant, note As Variant, loadedVals As Collection
    Set m_names = New Collection
    Set m_fields = New Collection
    Set m_vals = New Collection
    m_editForm = ""
    loaded = SameRecord(frm)
    If Not loaded And Not frm.NewRecord Then Exit Sub   ' no Form_Current note: a conflict keeps Access's window
    m_editForm = frm.Name
    If loaded Then note = m_recs(frm.Name): Set loadedVals = note(1)
    For Each c In frm.Controls
        If IsBoundEditor(c) Then
            src = c.ControlSource
            If loaded Then was = loadedVals(c.Name) Else was = Null     ' a new record: all the user's
            If Not SameValue(c.Value, was) Then
                m_names.Add c.Name
                m_fields.Add FieldCaption(c, src)
                m_vals.Add c.Value
            End If
        End If
    Next
End Sub

' True when KitNoteRecord noted frm's current record.
Private Function SameRecord(frm As Form) As Boolean
    Dim bm As Variant, note As Variant
    If m_recs Is Nothing Or frm.NewRecord Then Exit Function
    On Error GoTo Fail                  ' no note for frm
    note = m_recs(frm.Name)
    bm = frm.Bookmark
    SameRecord = (StrComp(bm, note(0), vbBinaryCompare) = 0)
    Exit Function
Fail:
End Function

' True for a control the user edits that is bound to a field.
Private Function IsBoundEditor(c As Control) As Boolean
    Dim src As String
    Select Case c.ControlType
        Case acTextBox, acComboBox, acCheckBox, acToggleButton, acListBox, acOptionGroup
            src = c.ControlSource
            IsBoundEditor = Len(src) > 0 And Left$(src, 1) <> "="
    End Select
End Function

' Tests only: answer the next conflicts with vbYes or vbNo instead of
' asking (0 asks again).
Public Sub SetConflictAnswer(ByVal answer As Long)
    m_testAnswer = answer
End Sub

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

' KitSaveRecord's work, retried while locked (KitSaveRecord keeps a second
' click out while it waits).
Private Function SaveTry(frm As Form) As Boolean
    Dim attempt As Long, num As Long
TryAgain:
    On Error GoTo Fail
    m_formLock = 0
    m_conflictForm = ""
    KitTakeRefused                      ' clear a note left by an earlier save
    m_saving = True
    If frm.Dirty Then frm.Dirty = False
    m_saving = False
    If Len(m_conflictForm) > 0 Then GoTo Conflict   ' Access gave up the save without an error
    SaveTry = Not frm.Dirty
    If SaveTry Then KitNoteRecord frm   ' the saved record is the new "as loaded" (AfterUpdate may not be wired)
    Exit Function
Conflict:
    ResolveConflict frm                 ' after the failed save, never inside it
    Exit Function
Fail:
    m_saving = False
    If Len(m_conflictForm) > 0 Then Resume Conflict ' a write conflict: the user decides
    num = Err.Number
    ' A form wired to KitFormError gets the lock in Form_Error first; it
    ' notes it and stays quiet, and the save fails here as refused (2101 etc.).
    If m_formLock <> 0 Then num = m_formLock
    m_formLock = 0
    If LockWait(num, attempt) Then Resume TryAgain
    If IsLockError(num) Then
        LogError "modKitLocks.KitSaveRecord", frm.Name & " locked after " & LOCK_TRIES & " retries"
        KitMsgBox LOCK_MSG, vbExclamation
        Exit Function
    End If
    If KitTakeRefused() Then Exit Function   ' BeforeUpdate refused and said why, whatever the number
    Select Case num
        Case 2101, 2115, 2169, 3020     ' refused by BeforeUpdate or Form_Error: it has said why
            Exit Function
    End Select
    RaiseAgain "modKitLocks.SaveTry", frm.Name
End Function

' After the failed save: drop the user's changes, read the record again
' (the other user's saved version), then ask. Yes puts the noted values back
' on top (the record is changed again, for them to save); No leaves the
' other user's version.
Private Sub ResolveConflict(frm As Form)
    Dim i As Long, list As String, answer As Long, names As Collection, vals As Collection, fields As Collection
    Set names = m_names: Set vals = m_vals: Set fields = m_fields
    m_conflictForm = ""
    m_editForm = ""
    frm.Undo
    ReloadRecord frm                ' Refresh alone leaves the record out of date (error 440 on the next edit)
    If names Is Nothing Then Exit Sub
    If names.Count = 0 Then Exit Sub
    For i = 1 To fields.Count
        If i <= 5 Then list = list & IIf(i > 1, ", ", "") & fields(i)
    Next
    If fields.Count > 5 Then list = list & " and " & (fields.Count - 5) & " more"
    answer = m_testAnswer
    If answer = 0 Then
        answer = KitMsgBox("Someone else saved this record while you were editing it." & vbCrLf & vbCrLf & _
            "You changed: " & list & "." & vbCrLf & vbCrLf & _
            "Keep your changes? Yes puts them back over the other person's, then you save again. " & _
            "No keeps the other person's version.", vbYesNo + vbExclamation, "Record changed by someone else")
    End If
    If answer <> vbYes Then Exit Sub
    For i = 1 To names.Count
        frm.Controls(names(i)).Value = vals(i)
    Next
End Sub

' Read frm's current record again: requery, then go back to it by its
' primary key (by position when the key can't be found).
Private Sub ReloadRecord(frm As Form)
    Dim crit As String, pos As Long, rs As DAO.Recordset
    pos = frm.CurrentRecord
    crit = KeyCriteria(frm)
    frm.Requery
    If Len(crit) > 0 Then
        Set rs = frm.RecordsetClone
        rs.FindFirst crit
        If Not rs.NoMatch Then frm.Bookmark = rs.Bookmark: Exit Sub
    End If
    If pos > 1 Then DoCmd.GoToRecord acDataForm, frm.Name, acGoTo, pos
End Sub

' A FindFirst criteria for frm's current record from the primary key of
' the table behind its first field ("[id]=12"), or "" when there is none.
Private Function KeyCriteria(frm As Form) As String
    Dim rs As DAO.Recordset, tbl As String, ix As DAO.Index, kf As DAO.Field, f As DAO.Field, crit As String, found As Boolean
    On Error GoTo Fail
    Set rs = frm.Recordset
    tbl = rs.Fields(0).SourceTable
    If Len(tbl) = 0 Then Exit Function
    For Each ix In CurrentDb.TableDefs(tbl).Indexes
        If ix.Primary Then
            For Each kf In ix.Fields
                found = False
                For Each f In rs.Fields
                    If f.SourceTable = tbl And StrComp(f.SourceField, kf.Name, vbTextCompare) = 0 Then
                        crit = crit & IIf(Len(crit) > 0, " AND ", "") & "[" & f.Name & "]=" & SqlLiteral(f.Value, f.Type)
                        found = True
                        Exit For
                    End If
                Next
                If Not found Then Exit Function     ' a key field the form doesn't load
            Next
            KeyCriteria = crit
            Exit Function
        End If
    Next
    Exit Function
Fail:
    KeyCriteria = ""                ' no key to go by: ReloadRecord uses the position
End Function

' v as a Jet SQL literal for a field of DAO type t.
Private Function SqlLiteral(ByVal v As Variant, ByVal t As Integer) As String
    If IsNull(v) Then
        SqlLiteral = "Null"
    Else
        Select Case t
            Case dbText, dbMemo, dbChar, dbGUID
                SqlLiteral = "'" & Replace(CStr(v), "'", "''") & "'"
            Case dbDate
                SqlLiteral = Format$(v, "\#mm\/dd\/yyyy hh\:nn\:ss\#")
            Case Else
                SqlLiteral = Str$(v)
        End Select
    End If
End Function

' What the user calls the field c shows: its attached label's caption
' without a trailing colon or "*", else the field name.
Private Function FieldCaption(c As Control, ByVal src As String) As String
    Dim cap As String
    On Error Resume Next                ' no attached label
    cap = c.Controls(0).Caption
    On Error GoTo 0
    cap = Trim$(Replace(cap, "&", ""))
    Do While Len(cap) > 0 And (Right$(cap, 1) = ":" Or Right$(cap, 1) = "*")
        cap = Trim$(Left$(cap, Len(cap) - 1))
    Loop
    If Len(cap) = 0 Then cap = Replace(Replace(src, "[", ""), "]", "")
    FieldCaption = cap
End Function

Private Function SameValue(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsNull(a) Or IsNull(b) Then
        SameValue = IsNull(a) And IsNull(b)
    Else
        SameValue = (StrComp(CStr(a), CStr(b), vbBinaryCompare) = 0)
    End If
End Function

' Wait ms milliseconds, letting Access paint.
Private Sub Pause(ByVal ms As Long)
    Dim t As Single
    t = Timer
    Do
        DoEvents
        Sleep 15
    Loop While (Timer - t) * 1000 < ms And Timer >= t
End Sub
