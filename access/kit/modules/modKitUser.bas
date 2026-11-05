Option Compare Database
Option Explicit

' Access UI kit - who is using the app, and who created or last changed each
' record.
'
' WinUser() is the Windows account name, read from Windows itself
' (GetUserNameW): the USERNAME environment variable can be changed by
' anyone before Access starts, so it is only the fallback. KitUser() is the
' name to record: the app's signed-in user (TempVars("Username"), set by an
' app with its own sign-in) or, without one, WinUser().
'
' Stamps: a table that several users edit gets four columns, added by
' AddStampFields at build or upgrade time:
'   created_by TEXT(150), created_at DATETIME, updated_by TEXT(150), updated_at DATETIME
' A bound form fills them from its BeforeUpdate:
'   Private Sub Form_BeforeUpdate(Cancel As Integer)
'       StampRecord Me
'   End Sub
' and code that writes with SQL adds StampSql():
'   KitExecute "UPDATE item SET status='Closed', " & StampSql() & " WHERE id=" & id
'   KitExecute "INSERT INTO item (name, " & StampColumns() & ") VALUES ('Pump', " & StampValues() & ")"

Public Const STAMP_CREATED_BY As String = "created_by"
Public Const STAMP_CREATED_AT As String = "created_at"
Public Const STAMP_UPDATED_BY As String = "updated_by"
Public Const STAMP_UPDATED_AT As String = "updated_at"
Public Const STAMP_USER_LEN As Long = 150      ' the by columns' width, as error_log.username

Private Declare PtrSafe Function GetUserNameW Lib "advapi32" (ByVal lpBuffer As LongPtr, nSize As Long) As Long
Private Declare PtrSafe Function GetComputerNameW Lib "kernel32" (ByVal lpBuffer As LongPtr, nSize As Long) As Long
Private Declare PtrSafe Function GetUserNameExW Lib "secur32" (ByVal NameFormat As Long, ByVal lpNameBuffer As LongPtr, nSize As Long) As Long
Private Declare PtrSafe Function GetCurrentProcess Lib "kernel32" () As LongPtr
Private Declare PtrSafe Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As LongPtr, ByVal DesiredAccess As Long, TokenHandle As LongPtr) As Long
Private Declare PtrSafe Function GetTokenInformation Lib "advapi32" (ByVal TokenHandle As LongPtr, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare PtrSafe Function LookupAccountSidW Lib "advapi32" (ByVal lpSystemName As LongPtr, ByVal Sid As LongPtr, ByVal Name As LongPtr, cchName As Long, ByVal ReferencedDomainName As LongPtr, cchReferencedDomainName As Long, peUse As Long) As Long
Private Declare PtrSafe Function ConvertSidToStringSidW Lib "advapi32" (ByVal Sid As LongPtr, StringSid As LongPtr) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function LocalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Const NAME_SAM_COMPATIBLE As Long = 2   ' EXTENDED_NAME_FORMAT: DOMAIN\user
Private Const TOKEN_QUERY As Long = 8
Private Const TOKEN_USER_CLASS As Long = 1      ' TOKEN_INFORMATION_CLASS.TokenUser

Private m_winUser As String
Private m_computer As String
Private m_account As String
Private m_sid As String

' The Windows account running Access ("jsmith"), read once from Windows; the
' USERNAME environment variable only if Windows doesn't answer.
Public Function WinUser() As String
    Dim buf As String, n As Long
    If Len(m_winUser) = 0 Then
        n = 257                                 ' UNLEN + 1
        buf = String$(n, vbNullChar)
        If GetUserNameW(StrPtr(buf), n) <> 0 And n > 1 Then m_winUser = Left$(buf, n - 1)
        If Len(m_winUser) = 0 Then m_winUser = Environ$("USERNAME")
    End If
    WinUser = m_winUser
End Function

' The computer's name ("NAVY-PC-042"), read once from Windows; the
' COMPUTERNAME environment variable only if Windows doesn't answer.
Public Function WinComputer() As String
    Dim buf As String, n As Long
    If Len(m_computer) = 0 Then
        n = 256
        buf = String$(n, vbNullChar)
        If GetComputerNameW(StrPtr(buf), n) <> 0 Then m_computer = Left$(buf, n)
        If Len(m_computer) = 0 Then m_computer = Environ$("COMPUTERNAME")
    End If
    WinComputer = m_computer
End Function

' The full Windows account, DOMAIN\user ("NAVY\jsmith", or "PC-042\jsmith"
' off a domain), read once from Windows: GetUserNameExW, else the account
' of Access's process token (LookupAccountSidW). "" if neither answers;
' never the USERDOMAIN variable, which anyone can set. Use this, not
' WinUser(), to match a person record: two domains can both have a jsmith.
Public Function WinAccount() As String
    Dim buf As String, n As Long, tok() As Byte, pSid As LongPtr, nm As String, dom As String, nNm As Long, nDom As Long, use As Long
    If Len(m_account) = 0 Then
        n = 512
        buf = String$(n, vbNullChar)
        If GetUserNameExW(NAME_SAM_COMPATIBLE, StrPtr(buf), n) <> 0 Then m_account = Left$(buf, n)
    End If
    If Len(m_account) = 0 Then                  ' e.g. no domain controller in reach
        pSid = TokenUserSid(tok)
        If pSid <> 0 Then
            nNm = 257: nDom = 256
            nm = String$(nNm, vbNullChar): dom = String$(nDom, vbNullChar)
            If LookupAccountSidW(0, pSid, StrPtr(nm), nNm, StrPtr(dom), nDom, use) <> 0 Then
                m_account = Left$(dom, nDom) & "\" & Left$(nm, nNm)
            End If
        End If
    End If
    WinAccount = m_account
End Function

' The Windows account's security ID ("S-1-5-21-...-1104"), from Access's own
' process token, read once. It stays the same when the account is renamed,
' so it is the surest way to know a person again. "" if Windows doesn't
' answer; there is no environment fallback, since a SID that can be typed
' in can't be trusted.
Public Function WinSid() As String
    Dim tok() As Byte, pSid As LongPtr, pStr As LongPtr, n As Long
    If Len(m_sid) > 0 Then WinSid = m_sid: Exit Function
    pSid = TokenUserSid(tok)
    If pSid = 0 Then Exit Function
    If ConvertSidToStringSidW(pSid, pStr) <> 0 Then
        n = lstrlenW(pStr)
        m_sid = String$(n, vbNullChar)
        CopyMemory ByVal StrPtr(m_sid), ByVal pStr, n * 2
        LocalFree pStr
    End If
    WinSid = m_sid
End Function

' The name to record as the person doing something: the app's signed-in
' user (TempVars("Username")) when it has one, else WinUser().
Public Function KitUser() As String
    Dim u As String
    On Error Resume Next                        ' TempVars may be unset
    u = Nz(TempVars("Username"), "")
    On Error GoTo 0
    If Len(u) = 0 Then u = WinUser()
    KitUser = Left$(u, STAMP_USER_LEN)
End Function

' A bound form's BeforeUpdate: fill the stamp fields its record source has.
' A new record gets created_by and created_at; every save sets updated_by
' and updated_at. Fields the record source lacks are skipped, so a form can
' call it before its table has all four.
Public Sub StampRecord(frm As Form)
    Dim t As Date, u As String
    t = Now()
    u = KitUser()
    If frm.NewRecord Then
        SetField frm, STAMP_CREATED_BY, u
        SetField frm, STAMP_CREATED_AT, t
    End If
    SetField frm, STAMP_UPDATED_BY, u
    SetField frm, STAMP_UPDATED_AT, t
End Sub

' For an UPDATE's SET list: "updated_by='jsmith', updated_at=Now()".
Public Function StampSql() As String
    StampSql = STAMP_UPDATED_BY & "=" & SqlText(KitUser()) & ", " & STAMP_UPDATED_AT & "=Now()"
End Function

' For an INSERT's column list, with StampValues: all four stamp columns.
Public Function StampColumns() As String
    StampColumns = STAMP_CREATED_BY & ", " & STAMP_CREATED_AT & ", " & STAMP_UPDATED_BY & ", " & STAMP_UPDATED_AT
End Function

' For an INSERT's VALUES list, matching StampColumns.
Public Function StampValues() As String
    Dim u As String
    u = SqlText(KitUser())
    StampValues = u & ", Now(), " & u & ", Now()"
End Function

' Build or upgrade step: add the stamp columns table lacks, in db (default
' the current database; pass the back end's to change a linked table). Only
' adds columns, never changes or removes one. Returns how many it added.
Public Function AddStampFields(ByVal table As String, Optional db As DAO.Database) As Long
    Dim cols As Variant, i As Long, td As DAO.TableDef
    On Error GoTo Fail
    If db Is Nothing Then Set db = CurrentDb
    db.TableDefs.Refresh                        ' a Database kept by the caller may not list columns added since
    cols = Array(STAMP_CREATED_BY & " TEXT(" & STAMP_USER_LEN & ")", STAMP_CREATED_AT & " DATETIME", _
        STAMP_UPDATED_BY & " TEXT(" & STAMP_USER_LEN & ")", STAMP_UPDATED_AT & " DATETIME")
    Set td = db.TableDefs(table)
    For i = 0 To UBound(cols)
        If Not HasField(td, Split(cols(i), " ")(0)) Then
            db.Execute "ALTER TABLE [" & table & "] ADD COLUMN " & cols(i), dbFailOnError
            AddStampFields = AddStampFields + 1
        End If
    Next
    Exit Function
Fail:
    RaiseAgain "modKitUser.AddStampFields", "table " & table
End Function

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

' The SID of the account running Access, from its process token: a
' pointer into tok (TOKEN_USER), so tok must outlive its use. 0 if Windows
' doesn't answer.
Private Function TokenUserSid(tok() As Byte) As LongPtr
    Dim hTok As LongPtr, need As Long, p As LongPtr
    If OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, hTok) = 0 Then Exit Function
    GetTokenInformation hTok, TOKEN_USER_CLASS, ByVal 0&, 0, need     ' asks the size; fails by design
    If need > 0 Then
        ReDim tok(0 To need - 1)
        If GetTokenInformation(hTok, TOKEN_USER_CLASS, tok(0), need, need) <> 0 Then
            CopyMemory p, tok(0), LenB(p)                          ' TOKEN_USER starts with the SID pointer
        End If
    End If
    CloseHandle hTok
    TokenUserSid = p
End Function

' Set field nm of frm's current record, when its record source has it.
Private Sub SetField(frm As Form, ByVal nm As String, ByVal v As Variant)
    If Not HasField(frm.Recordset, nm) Then Exit Sub
    frm(nm).Value = v                           ' the record being saved, not the recordset's
End Sub

' True when obj (a TableDef or Recordset) has field nm.
Private Function HasField(obj As Object, ByVal nm As String) As Boolean
    Dim fld As DAO.Field
    For Each fld In obj.Fields
        If StrComp(fld.Name, nm, vbTextCompare) = 0 Then HasField = True: Exit Function
    Next
End Function

' Text as a quoted SQL literal.
Private Function SqlText(ByVal s As String) As String
    SqlText = "'" & Replace(s, "'", "''") & "'"
End Function
