Option Compare Database
Option Explicit

' Access UI kit - finding a person's Windows account in Active Directory, so
' Personnel Import can link them before they ever open an app.
'
' PersonnelFindAccount takes what a people list has, a user principal name
' (UPN, "1234567890.A@smil.mil"), a DoD ID (made into a UPN with
' PERSONNEL_UPN_SUFFIX) or an email address, and returns the account's
' DOMAIN\user and SID. A UPN or DOMAIN\user goes straight to Windows
' (LookupAccountNameW, which takes UPNs and searches every domain in the
' forest); an email address is first looked up in the global catalog (mail
' or proxyAddresses) for its UPN. Nothing here needs RSAT or PowerShell: it
' is what Get-ADUser -Filter "UserPrincipalName -eq '...'" | Select SID does,
' with Windows' own calls, on any domain-joined PC.
'
' Nothing here saves anything. modPersonnel keeps a person's UPN, and so
' their DoD ID, in kit_person_upn, which only admins can read and which is
' never exported or logged.

' A DoD ID (10 digits) becomes this UPN: <DoD ID> & PERSONNEL_UPN_SUFFIX.
Public Const PERSONNEL_UPN_SUFFIX As String = ".A@smil.mil"

Private Const SID_TYPE_USER As Long = 1
Private Const MAX_SID_BYTES As Long = 68

Private Declare PtrSafe Function LookupAccountNameW Lib "advapi32" (ByVal lpSystemName As LongPtr, ByVal lpAccountName As LongPtr, _
    ByVal Sid As LongPtr, cbSid As Long, ByVal ReferencedDomainName As LongPtr, cchReferencedDomainName As Long, peUse As Long) As Long
Private Declare PtrSafe Function LookupAccountSidW Lib "advapi32" (ByVal lpSystemName As LongPtr, ByVal Sid As LongPtr, _
    ByVal Name As LongPtr, cchName As Long, ByVal ReferencedDomainName As LongPtr, cchReferencedDomainName As Long, peUse As Long) As Long
Private Declare PtrSafe Function ConvertSidToStringSidW Lib "advapi32" (ByVal Sid As LongPtr, StringSid As LongPtr) As Long
Private Declare PtrSafe Function LocalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private m_gcBase As String          ' "GC://DC=..." once found; "-" when there is no directory to ask

' The UPN for a DoD ID ("1234567890" gives "1234567890.A@smil.mil"), or ""
' when dodId isn't 10 digits.
Public Function PersonnelUpnFor(ByVal dodId As String) As String
    dodId = Trim$(dodId)
    If dodId Like "##########" Then PersonnelUpnFor = dodId & PERSONNEL_UPN_SUFFIX
End Function

' Find the user account for who: a UPN, DOMAIN\user, or an email address.
' True with account (DOMAIN\user) and sid ("S-1-5-21-...") set when exactly
' that user account is found; False, with both "", when it isn't, when it is
' a group or computer, or when no directory answers. Never raises.
Public Function PersonnelFindAccount(ByVal who As String, account As String, sid As String) As Boolean
    Dim upn As String
    On Error GoTo Fail
    account = "": sid = ""
    who = Trim$(who)
    If Len(who) = 0 Then Exit Function
    upn = who
    If InStr(who, "@") > 0 And InStr(who, "\") = 0 Then
        ' A UPN first; an address that isn't one is looked up as email.
        If AccountFromName(upn, account, sid) Then PersonnelFindAccount = True: Exit Function
        upn = UpnForEmail(who)
        If Len(upn) = 0 Then Exit Function
    End If
    PersonnelFindAccount = AccountFromName(upn, account, sid)
    Exit Function
Fail:
    account = "": sid = ""
End Function

' LookupAccountNameW for who, then LookupAccountSidW for the account's
' DOMAIN\user (the form WinAccount() gives). Only user accounts count.
Private Function AccountFromName(ByVal who As String, account As String, sid As String) As Boolean
    Dim sidBuf() As Byte, cbSid As Long, dom As String, nDom As Long, use As Long
    Dim nm As String, nNm As Long, pStr As LongPtr, n As Long
    ReDim sidBuf(0 To MAX_SID_BYTES - 1)
    cbSid = MAX_SID_BYTES
    nDom = 256: dom = String$(nDom, vbNullChar)
    If LookupAccountNameW(0, StrPtr(who), VarPtr(sidBuf(0)), cbSid, StrPtr(dom), nDom, use) = 0 Then Exit Function
    If use <> SID_TYPE_USER Then Exit Function
    nNm = 257: nm = String$(nNm, vbNullChar)
    nDom = 256: dom = String$(nDom, vbNullChar)
    If LookupAccountSidW(0, VarPtr(sidBuf(0)), StrPtr(nm), nNm, StrPtr(dom), nDom, use) = 0 Then Exit Function
    If ConvertSidToStringSidW(VarPtr(sidBuf(0)), pStr) = 0 Then Exit Function
    n = lstrlenW(pStr)
    sid = String$(n, vbNullChar)
    CopyMemory ByVal StrPtr(sid), ByVal pStr, n * 2
    LocalFree pStr
    account = Left$(dom, nDom) & "\" & Left$(nm, nNm)
    AccountFromName = True
End Function

' The UPN of the one user whose mail or proxy address is email, from the
' forest's global catalog; "" when there is none, more than one, or no
' directory to ask.
Private Function UpnForEmail(ByVal email As String) As String
    Dim cn As Object, rs As Object, f As String
    On Error GoTo Fail
    If Len(m_gcBase) = 0 Then
        m_gcBase = "-"
        m_gcBase = "GC://" & GetObject("LDAP://RootDSE").Get("rootDomainNamingContext")
    End If
    If m_gcBase = "-" Then Exit Function
    f = LdapText(email)
    Set cn = CreateObject("ADODB.Connection")
    cn.Provider = "ADsDSOObject"
    cn.Open "Active Directory Provider"
    Set rs = cn.Execute("<" & m_gcBase & ">;(&(objectCategory=person)(objectClass=user)(|(mail=" & f & ")(proxyAddresses=smtp:" & f & ")));" & _
        "userPrincipalName;subtree")
    If Not rs.EOF Then
        UpnForEmail = Nz(rs.Fields(0).Value, "")
        rs.MoveNext
        If Not rs.EOF Then UpnForEmail = ""     ' two people share it: not proof of either
    End If
    rs.Close
    cn.Close
    Exit Function
Fail:
    UpnForEmail = ""
End Function

' s with the characters an LDAP filter treats as special escaped (RFC 4515).
Private Function LdapText(ByVal s As String) As String
    s = Replace(s, "\", "\5c")
    s = Replace(s, "*", "\2a")
    s = Replace(s, "(", "\28")
    s = Replace(s, ")", "\29")
    LdapText = Replace(s, vbNullChar, "\00")
End Function
