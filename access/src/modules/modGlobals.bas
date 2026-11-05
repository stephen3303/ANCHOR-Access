Option Compare Database
Option Explicit

' ANCHOR (Access edition): the ANCHOR message-handling web app rebuilt as an
' Access app on the access-modern-ui kit. App constants, build-time silence and the
' small helpers every module uses: Notify (messages), Q (SQL text), NavTo,
' and at the end FieldText, Plural, TableExists, InsertRow and
' RollbackKeepingError.
' The names the kit calls on the app (README, "What the app must provide")
' are here and in modUI; MeasureText comes from the kit's modTextMeasure
' adapter (kit\adapters), which build.ps1 loads with the kit.
'
' Owner: the app core (build phase 1). Depends on: modErrors; the kit's
' modDarkMode, modKitLocks, modKitShell, modKitUI, modKitUtil.

Public Const APP_NAME As String = "ANCHOR"
Public Const APP_CODE As String = "ANCHOR"      ' this app's code in the shared Personnel file (kit_person_app)
Public Const BE_FILE As String = "ANCHOR_BE.accdb"   ' the shared data file's usual name
Public g_Silent As Boolean                       ' True during the build: nothing may block

' InsertRow's words for a lock, the same as KitExecute's (modKitLocks keeps
' its own Private).
Private Const INSERT_BUSY As String = "Waiting for another user to finish saving..."
Private Const INSERT_LOCKED As String = "Another user is saving the same records right now. Wait a moment, then try again."

' Message to the user. During the build it is printed instead; warnings and
' errors raise, and questions are answered Yes.
Public Function Notify(ByVal msg As String, Optional ByVal style As VbMsgBoxStyle = vbInformation) As VbMsgBoxResult
    If g_Silent Then
        Debug.Print msg
        If (style And vbExclamation) = vbExclamation Or (style And vbCritical) = vbCritical Then
            Err.Raise vbObjectError + 999, APP_NAME, msg
        End If
        If (style And vbYesNo) = vbYesNo Then Notify = vbYes Else Notify = vbOK
    Else
        Notify = KitMsgBox(msg, style, APP_NAME)    ' centred over Access (modKitUI)
    End If
End Function

' Text as a quoted SQL string literal, its single quotes doubled: Q("O'Neil")
' gives 'O''Neil'. Use it for every value put into SQL as text.
'   s   any value; Null gives '' (an empty string, not the SQL Null)
' Errors: none.
Public Function Q(ByVal s As Variant) As String
    Q = "'" & Replace(Nz(s, ""), "'", "''") & "'"
End Function

' Text for a LIKE pattern typed by the user: its wildcard characters are
' bracketed so they match themselves, as GridHighlight marks them.
Public Function LikeLiteral(ByVal s As String) As String
    s = Replace(s, "[", "[[]")
    LikeLiteral = Replace(Replace(Replace(s, "*", "[*]"), "?", "[?]"), "#", "[#]")
End Function

' Go to a page once the calling event has finished (the kit's ShellGo).
Public Function NavTo(ByVal key As String)
    NavTo = ShellGo(key)
End Function

' Open the palette picker, only while the current page can be left, because
' a palette switch reopens the main window (frmMain.CanLeavePage).
Public Function AppAppearance()
    On Error GoTo Err_Handler
    If IsFormOpen("frmMain") Then
        If Not Forms("frmMain").CanLeavePage() Then Exit Function
    End If
    ShowAppearance
    Exit Function
Err_Handler:
    HandleError "modGlobals.AppAppearance"
End Function

' A text box's text: what is typed so far while it has the focus (Ctrl+S
' saves without leaving it), else its value. frm is the form it is on.
Public Function FieldText(frm As Form, c As Control) As String
    Dim act As String
    On Error Resume Next
    act = frm.ActiveControl.Name            ' error 2474 when none has the focus
    On Error GoTo 0
    If act = c.Name Then FieldText = Nz(c.Text, "") Else FieldText = Nz(c.Value, "")
End Function

' word, with an s added unless n is 1: "1 message", "3 messages".
Public Function Plural(ByVal n As Long, ByVal word As String) As String
    Plural = word & IIf(n = 1, "", "s")
End Function

' True when this front end has a table, or a link to one, called nm (any
' case). localOnly: only a table of its own, not a link.
Public Function TableExists(ByVal nm As String, Optional ByVal localOnly As Boolean = False) As Boolean
    Dim td As DAO.TableDef
    For Each td In CurrentDb.TableDefs
        If StrComp(td.Name, nm, vbTextCompare) = 0 Then
            TableExists = (Not localOnly) Or Len(td.Connect) = 0
            Exit Function
        End If
    Next
End Function

' Run one INSERT that adds one row and return the row's AutoNumber id. The
' id is read with SELECT @@IDENTITY on the Database object that ran the
' insert, which is the only object that sees it (each CurrentDb is a new
' one), so a new row is found without a look-up by name. Inside a
' DBEngine.Workspaces(0) transaction the insert is part of it.
'   sql   the INSERT ... VALUES or INSERT ... SELECT statement
' Returns the new id.
' Errors: 5 when the statement added no row or more than one; the engine's
'   error when the insert fails, which runs it only once. A lock is retried
'   as KitExecute retries it (the kit's LockWait, under the busy notice);
'   one that doesn't clear is logged and raised with the kit's plain
'   sentence, as KitExecute raises it.
Public Function InsertRow(ByVal sql As String) As Long
    Dim db As DAO.Database, rs As DAO.Recordset, attempt As Long, num As Long, d As String, src As String, busy As Boolean
    Set db = CurrentDb
TryAgain:
    On Error GoTo InsertFailed
    db.Execute sql, dbFailOnError
    On Error GoTo Fail
    If busy Then KitBusyEnd: busy = False
    If db.RecordsAffected <> 1 Then Err.Raise 5, "modGlobals.InsertRow", "Expected to add one row, added " & db.RecordsAffected & "."
    Set rs = db.OpenRecordset("SELECT @@IDENTITY", dbOpenSnapshot)
    InsertRow = Nz(rs.Fields(0).Value, 0)
    rs.Close
    Exit Function
InsertFailed:
    ' The insert failed. KitBusy and KitBusyEnd reset Err, so keep the error
    ' to pass it on.
    num = Err.Number: d = Err.Description: src = Err.Source
    If attempt = 0 And IsLockError(num) Then KitBusy INSERT_BUSY: busy = True
    If LockWait(num, attempt) Then Resume TryAgain
    If busy Then KitBusyEnd
    Err.Number = num: Err.Description = d: Err.Source = src
    If IsLockError(num) Then
        LogError "modGlobals.InsertRow", "locked after " & LOCK_TRIES & " retries: " & Left$(sql, 200)
        On Error GoTo 0
        Err.Raise num, "modGlobals.InsertRow", INSERT_LOCKED
    End If
    RaiseAgain "modGlobals.InsertRow", Left$(sql, 200)      ' not a lock: passed on without running it again
Fail:
    RaiseAgain "modGlobals.InsertRow", Left$(sql, 200)
End Function

' Roll back ws's transaction from an error handler and leave the error being
' handled as it was, so the handler can still pass it on (RaiseAgain) or show
' it (HandleError). A failed rollback is ignored: the error that caused it is
' the one to report.
'   ws   the workspace whose BeginTrans is open, DBEngine.Workspaces(0)
Public Sub RollbackKeepingError(ws As DAO.Workspace)
    Dim num As Long, desc As String, src As String
    num = Err.Number: desc = Err.Description: src = Err.Source
    On Error Resume Next
    ws.Rollback
    On Error GoTo 0
    Err.Number = num: Err.Description = desc: Err.Source = src
End Sub
