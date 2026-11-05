Option Compare Database
Option Explicit

' Access UI kit - trusted-folder check at startup. When the app's folder is
' not a Trusted Location, Access asks the user to enable content every time
' the app opens (or opens it with its code turned off). BuildTrustCheck gives
' the app a startup that notices and shows dlgTrust, a plain page explaining
' how to add the folder in the Trust Center. Nothing here changes Trust Center
' settings: the registry is only read.
'
' How startup works (BuildTrustCheck sets it up):
'   AutoExec macro   runs instead of a StartUpForm, because a macro runs even
'                    when the code is turned off.
'     code off       (CurrentProject.IsTrusted False): opens dlgTrust and stops.
'     code on        runs TrustStartup: opens the main form, then dlgTrust over
'                    it when Access will ask again next time: the folder isn't
'                    a Trusted Location and the user chose Enable Content only
'                    for this session. Once Access remembers the file as a
'                    trusted document, or every macro is allowed, the page
'                    stays away.
'   dlgTrust         has no code, no Open hook and no expressions, so it works
'                    with the code off. It isn't modal, so Access's Enable
'                    Content bar stays usable while it is open. Its Close button
'                    runs the macro mcrTrustQuit (quit Access: with the code off
'                    the app can't run); ShowTrustHelp, with the code on, points it
'                    at mcrTrustClose (close the page) and fills in the folder.

Private Const TRUST_FORM As String = "dlgTrust"
Private Const CLOSE_MACRO As String = "mcrTrustClose"
Private Const QUIT_MACRO As String = "mcrTrustQuit"
Private Const MAIN_FORM_PROP As String = "KitMainForm"    ' database property: the form TrustStartup opens

Private Const HKEY_CURRENT_USER As LongPtr = &H80000001
Private Const KEY_READ As Long = &H20019
Private Const ERROR_SUCCESS As Long = 0
Private Const ERROR_MORE_DATA As Long = 234
Private Const REG_DWORD As Long = 4
Private Const DRIVE_REMOTE As Long = 4

Private Declare PtrSafe Function RegOpenKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal subKey As LongPtr, _
    ByVal options As Long, ByVal samDesired As Long, phkResult As LongPtr) As Long
Private Declare PtrSafe Function RegEnumKeyExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal index As Long, _
    ByVal name As LongPtr, cchName As Long, ByVal reserved As LongPtr, ByVal cls As LongPtr, _
    ByVal cchClass As LongPtr, ByVal lastWrite As LongPtr) As Long
Private Declare PtrSafe Function RegQueryValueExW Lib "advapi32" (ByVal hKey As LongPtr, ByVal valueName As LongPtr, _
    ByVal reserved As LongPtr, valueType As Long, ByVal data As LongPtr, cbData As Long) As Long
Private Declare PtrSafe Function RegEnumValueW Lib "advapi32" (ByVal hKey As LongPtr, ByVal index As Long, _
    ByVal valueName As LongPtr, cchValueName As Long, ByVal reserved As LongPtr, ByVal valueType As LongPtr, _
    ByVal data As LongPtr, ByVal cbData As LongPtr) As Long
Private Declare PtrSafe Function RegCloseKey Lib "advapi32" (ByVal hKey As LongPtr) As Long
Private Declare PtrSafe Function ExpandEnvironmentStringsW Lib "kernel32" (ByVal src As LongPtr, _
    ByVal dst As LongPtr, ByVal size As Long) As Long
Private Declare PtrSafe Function GetDriveTypeW Lib "kernel32" (ByVal root As LongPtr) As Long

' ---------------------------------------------------------------------------
' Build
' ---------------------------------------------------------------------------

' Build step, in place of setting StartUpForm: builds dlgTrust and the macros
' AutoExec, mcrTrustClose and mcrTrustQuit, and remembers mainForm for
' TrustStartup.
' mainForm is the form the app starts on: the login form, frmMain, or the
' splash screen (frmSplash) when the app has one. Call it after anything that
' sets StartUpForm: it removes StartUpForm, so the first form opens only
' through AutoExec (with the code off it doesn't open at all). It replaces any
' AutoExec macro the app already has.
Public Sub BuildTrustCheck(ByVal mainForm As String)
    Dim db As DAO.Database, n As Long, d As String
    Build_dlgTrust
    ' Save stays at Prompt (0): with the code off Access refuses Close with any other
    ' Save argument. dlgTrust is never changed, so no save prompt can appear.
    LoadMacro CLOSE_MACRO, MacroRow("", "Close", "2", TRUST_FORM, "0")      ' form dlgTrust, Save Prompt
    ' Quit (QuitAccess in the macro builder) with Options at Prompt, for the same reason.
    LoadMacro QUIT_MACRO, MacroRow("", "Quit", "0")
    LoadMacro "AutoExec", _
        MacroRow("Not [CurrentProject].[IsTrusted]", "OpenForm", TRUST_FORM, "0", "", "", "-1", "0") & _
        MacroRow("Not [CurrentProject].[IsTrusted]", "StopMacro") & _
        MacroRow("", "RunCode", "TrustStartup()")
    Set db = CurrentDb
    On Error Resume Next
    db.Properties(MAIN_FORM_PROP) = mainForm
    If Err.Number = 3270 Then           ' property not found: create it
        Err.Clear
        db.Properties.Append db.CreateProperty(MAIN_FORM_PROP, dbText, mainForm)
    End If
    If Err.Number <> 0 Then
        n = Err.Number: d = Err.Description
        On Error GoTo 0
        Err.Raise n, , "Could not save the main form name: " & d
    End If
    Err.Clear
    db.Properties.Delete "StartUpForm"  ' error 3265 when there is none
    On Error GoTo 0
End Sub

' dlgTrust: heading, why it shows, Enable Content, the folder (selectable, to
' copy), the Trust Center steps, and a note for computers where IT manages the
' settings. Text heights come from MeasureText, so each block sits
' Sp(SP_ROW_GAP) under the one above. Built with no code module, no Open hook
' and no expressions (it must open with the code turned off); ShowTrustHelp
' themes it and shows the folder when the code is on. With the code off the
' folder box can't be filled, so a hint in its place says where Access shows it.
Private Sub Build_dlgTrust()
    Dim c As Control, pad As Long, w As Long, x As Long, y As Long, h As Long, s As String, boxH As Long
    Const DLG_W As Long = 7200          ' 480 px
    pad = Sp(SP_DIALOG)
    x = pad: w = DLG_W - 2 * pad
    NewDialog APP_NAME, DLG_W, px(400)
    F.HasModule = False
    F.Modal = False                     ' Access's Enable Content bar must work while the page is open
    F.CloseButton = False               ' with the code off the window's X would leave an empty Access open
    y = pad
    Lbl "lblTitle", "Trust this app's folder", x, y, w, px(26), 13, True
    y = y + px(26) + Sp(SP_ROW_GAP)
    s = "Access doesn't trust the folder this app is in, so it asks you to enable content, or turns the " & _
        "app's code off, each time the app opens. Trust the folder once and the app opens without asking."
    h = TextH(s, 10, w)
    Lbl "lblWhy", s, x, y, w, h
    y = y + h + Sp(SP_ROW_GAP)
    s = "If Access shows a bar with Enable Content at the top, choosing it turns the app's code on now. " & _
        "Access usually remembers that for this file, and then this page doesn't show again. If it still " & _
        "shows next time, trust the folder."
    h = TextH(s, 10, w)
    Lbl "lblEnable", s, x, y, w, h
    y = y + h + Sp(SP_ROW_GAP) * 2
    Lbl "lblFolderCap", "FOLDER", x, y, w, px(16), 8, True, CLR_MUTED_FG
    y = y + px(18)
    s = "In Access, choose File, then Info: the folder is in the path shown under the app's name."
    boxH = TextH(s, 10, w - px(16)) + px(8)
    If boxH < px(44) Then boxH = px(44)
    Set c = RO("txtFolder", "", x, y, w, boxH, 10)     ' ShowTrustHelp fills it: no expression
    FolderBox c
    c.Visible = False
    Set c = Lbl("lblFolderHint", s, x, y, w, boxH)     ' shown in its place with the code off
    FolderBox c
    y = y + boxH + Sp(SP_ROW_GAP) * 2
    Lbl "lblStepsCap", "TO TRUST IT", x, y, w, px(16), 8, True, CLR_MUTED_FG
    y = y + px(18)
    s = "1. In Access, choose File, then Options, then Trust Center, then Trust Center Settings." & vbCrLf & _
        "2. Choose Trusted Locations, then Add new location." & vbCrLf & _
        "3. Browse to the folder above and choose OK." & vbCrLf & _
        "4. If the folder is on a network drive, first tick Allow Trusted Locations on my network." & vbCrLf & _
        "5. Close Access and open the app again."
    h = TextH(s, 10, w)
    Lbl "lblSteps", s, x, y, w, h
    y = y + h + Sp(SP_ROW_GAP)
    s = "If these settings are locked on your computer, ask your IT support to add the folder."
    h = TextH(s, 9, w)
    Lbl "lblIT", s, x, y, w, h, 9, False, CLR_MUTED_FG
    y = y + h + Sp(SP_DIALOG)
    Set c = Btn("btnClose", "Close", DLG_W - pad - px(96), y, px(96), px(32))
    c.Default = True: c.Cancel = True
    F.Section(acDetail).Height = y + px(32) + pad
    ClearEvents F
    c.OnClick = QUIT_MACRO              ' a macro, not code: works with the code off
    SaveForm TRUST_FORM
End Sub

' The folder box's look, for txtFolder and the hint in its place.
Private Sub FolderBox(c As Control)
    c.BackStyle = 1: c.BackColor = CLR_SURFACE
    c.BorderStyle = 1: c.BorderColor = CLR_BORDER
    c.LeftMargin = px(8): c.TopMargin = px(4): c.RightMargin = px(8)
End Sub

' Empty every event property (On...) of frm and its controls: whatever
' NewDialog or the control builders hook up now or later (ThemeForm on Open,
' a backdrop on Close, [Event Procedure] and a focus-ring expression on
' buttons) would fail with the code off. Properties a control doesn't take
' are skipped.
Private Sub ClearEvents(frm As Object)
    Dim c As Control
    ClearOnProps frm.Properties
    For Each c In frm.Controls
        ClearOnProps c.Properties
    Next
End Sub

Private Sub ClearOnProps(props As Object)
    Dim prp As Object
    On Error Resume Next
    For Each prp In props
        If Left$(prp.Name, 2) = "On" Then
            If Len(Nz(prp.Value, "")) > 0 Then prp.Value = ""
        End If
    Next
    On Error GoTo 0
End Sub

' Height in twips that s takes in a label w wide at size points: each line
' (split at vbCrLf) wraps to as many lines as its measured width needs, with
' 10% spare for word breaks, as dlgMessage measures its message.
Private Function TextH(ByVal s As String, ByVal size As Integer, ByVal w As Long) As Long
    Dim para As Variant, lines As Long
    For Each para In Split(s, vbCrLf)
        If Len(para) = 0 Then
            lines = lines + 1
        Else
            lines = lines - Int(-(MeasureText(CStr(para), size, 400) * 1.1) / w)
        End If
    Next
    If lines < 1 Then lines = 1
    TextH = lines * CLng(size * 20 * 1.36) + 30
End Function

' One row of a macro in Access's text format (LoadFromText acMacro), with an
' optional condition.
Private Function MacroRow(ByVal cond As String, ByVal action As String, ParamArray args() As Variant) As String
    Dim s As String, i As Long
    s = "Begin" & vbCrLf
    If Len(cond) > 0 Then s = s & "    Condition =""" & cond & """" & vbCrLf
    s = s & "    Action =""" & action & """" & vbCrLf
    For i = 0 To UBound(args)
        s = s & "    Argument =""" & args(i) & """" & vbCrLf
    Next
    MacroRow = s & "End" & vbCrLf
End Function

' Replace macro name with the given rows (a temporary file, as LoadFromText
' needs one).
Private Sub LoadMacro(ByVal name As String, ByVal rows As String)
    Dim path As String, fh As Integer, n As Long, d As String
    path = Environ$("TEMP") & "\kit_macro_" & name & ".txt"
    fh = FreeFile
    Open path For Output As #fh
    Print #fh, "Version =196611" & vbCrLf & "ColumnsShown =2" & vbCrLf & rows;
    Close #fh
    On Error Resume Next
    DoCmd.DeleteObject acMacro, name
    Err.Clear
    Application.LoadFromText acMacro, name, path
    n = Err.Number: d = Err.Description
    Kill path                           ' the temporary file goes whether or not the load worked
    On Error GoTo 0
    If n <> 0 Then Err.Raise n, , "Could not load the macro " & name & ": " & d
End Sub

' ---------------------------------------------------------------------------
' Run time
' ---------------------------------------------------------------------------

' AutoExec, with the code on: open the main form, then the trust page over it
' when Access will ask to enable content again next time (TrustNeeded). Does
' nothing under automation (a build script opening the file), so nothing
' opens or waits there.
Public Function TrustStartup()
    Dim mainForm As String
    On Error GoTo Err_Handler
    If Not Application.UserControl Then Exit Function
    On Error Resume Next
    mainForm = CurrentDb.Properties(MAIN_FORM_PROP)
    On Error GoTo Err_Handler
    If Len(mainForm) > 0 Then DoCmd.OpenForm mainForm
    If TrustNeeded() Then ShowTrustHelp
    Exit Function
Err_Handler:
    HandleError "modTrust.TrustStartup", "main form " & mainForm
End Function

' True when Access will open this file with a prompt (or the code off) next
' time: not in a Trusted Location, not remembered as a trusted document (the
' user chose Enable Content and Access kept it, under TrustRecords, unless
' trusted documents are turned off), and macros not all enabled
' (VBAWarnings = 1, the policy's value winning over the user's). Read only.
Public Function TrustNeeded() As Boolean
    Dim sec As String, pol As String
    On Error GoTo Fail
    sec = "Software\Microsoft\Office\" & Application.Version & "\Access\Security"
    pol = "Software\Policies\Microsoft\Office\" & Application.Version & "\Access\Security"
    If RegDword(pol, "", "VBAWarnings", RegDword(sec, "", "VBAWarnings", 0)) = 1 Then Exit Function
    If InTrustedLocation() Then Exit Function
    ' Trusted documents count unless policy (or, without one, the user) turns them off.
    If RegDword(pol, "Trusted Documents", "DisableTrustedDocuments", _
            RegDword(sec, "Trusted Documents", "DisableTrustedDocuments", 0)) <> 1 Then
        ' Office 2010 and later keep them under Trusted Documents\TrustRecords;
        ' Trusted Records is read too, in case a version keeps them there.
        If TrustedRecord(sec & "\Trusted Documents\TrustRecords", CurrentProject.FullName) Then Exit Function
        If TrustedRecord(sec & "\Trusted Records", CurrentProject.FullName) Then Exit Function
    End If
    TrustNeeded = True
    Exit Function
Fail:
    LogError "modTrust.TrustNeeded"
    TrustNeeded = True
End Function

' True when the key root (TrustRecords) has a value named for file. Office
' writes the names in more than one form (C:\..., file:///C:/..., with
' %USERPROFILE% for the profile folder), so each is normalized before comparing.
Private Function TrustedRecord(ByVal root As String, ByVal file As String) As Boolean
    Dim hKey As LongPtr, i As Long, nm As String, n As Long, rc As Long
    file = NormPath(file)
    If RegOpenKeyExW(HKEY_CURRENT_USER, StrPtr(root), 0, KEY_READ, hKey) <> ERROR_SUCCESS Then Exit Function
    Do
        nm = String$(2048, vbNullChar): n = 2048
        rc = RegEnumValueW(hKey, i, StrPtr(nm), n, 0, 0, 0, 0)
        If rc = ERROR_SUCCESS Then
            If RecordPath(Left$(nm, n)) = file Then TrustedRecord = True: Exit Do
        ElseIf rc <> ERROR_MORE_DATA Then
            Exit Do                     ' ERROR_NO_MORE_ITEMS; a name too long to be this file is skipped
        End If
        i = i + 1
    Loop
    RegCloseKey hKey
End Function

' A TrustRecords value name as a normalized Windows path. Access writes them
' URL-style: %USERPROFILE%/Projects/Navy%20Access/app.accdb. Environment
' variables are expanded first, so a name like %DATA% is never read as an
' escape; then %XX escapes are decoded. Public so the kit's smoke test can
' check it.
Public Function RecordPath(ByVal s As String) As String
    If LCase$(Left$(s, 8)) = "file:///" Then
        s = Mid$(s, 9)
    ElseIf LCase$(Left$(s, 7)) = "file://" Then
        s = "\\" & Mid$(s, 8)
    End If
    RecordPath = NormPath(UrlDecode(ExpandEnv(Replace(s, "/", "\"))))
End Function

' %XX escapes decoded where XX is hex for an ASCII character (a space is
' %20). Other escapes (UTF-8 bytes of non-ASCII names) stay as they are, so
' such a file never matches and the trust page shows: the safe side.
Private Function UrlDecode(ByVal s As String) As String
    Dim i As Long, out As String, hx As String
    i = 1
    Do While i <= Len(s)
        hx = Mid$(s, i + 1, 2)
        If Mid$(s, i, 1) = "%" And Len(hx) = 2 And hx Like "[0-7][0-9A-Fa-f]" Then
            out = out & Chr$(CLng("&H" & hx))
            i = i + 3
        Else
            out = out & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    UrlDecode = out
End Function

' Environment variables in p expanded (%USERPROFILE% and so on).
Private Function ExpandEnv(ByVal p As String) As String
    Dim buf As String, n As Long
    If InStr(p, "%") > 0 Then
        buf = String$(1024, vbNullChar)
        n = ExpandEnvironmentStringsW(StrPtr(p), StrPtr(buf), 1024)
        If n > 0 And n <= 1024 Then p = Left$(buf, n - 1)
    End If
    ExpandEnv = p
End Function

' Open the trust page (dlgTrust) in the current palette, centred over Access
' like the other dialogs. It has no code of its own (AutoCenter places it when
' the code is off), so it opens hidden, is themed and placed, then shown. With
' the code on, its Close only closes the page and the folder box shows the
' folder (neither is saved: the page opens as built next time).
Public Function ShowTrustHelp()
    On Error GoTo Err_Handler
    DoCmd.OpenForm TRUST_FORM, WindowMode:=acHidden
    With Forms(TRUST_FORM)
        .Controls("btnClose").OnClick = CLOSE_MACRO
        .Controls("txtFolder").Value = CurrentProject.Path
        .Controls("txtFolder").Visible = True
        .Controls("lblFolderHint").Visible = False
    End With
    ThemeForm Forms(TRUST_FORM)
    CenterOnApp Forms(TRUST_FORM)
    Forms(TRUST_FORM).Visible = True
    Exit Function
Err_Handler:
    HandleError "modTrust.ShowTrustHelp"
End Function

' True when folder (default: the app's own) is inside a Trusted Location
' Access will honor: listed in the policy Trusted Locations, or the user's own
' when policy allows them,
' with "subfolders" ticked when folder is below it, network locations allowed
' when it is on a network drive, and trusted locations not turned off by
' policy. False if the settings can't be read, so the worst case is showing
' the trust page when it wasn't needed.
Public Function InTrustedLocation(Optional ByVal folder As String = "") As Boolean
    Dim pol As String, common As String, usr As String, allowNet As Boolean, isNet As Boolean, onList As Boolean
    On Error GoTo Fail
    If Len(folder) = 0 Then folder = CurrentProject.Path
    folder = NormPath(folder)
    ' Policy locations for Access, policy locations for every Office app
    ' (Security Settings\Trust Center), and the user's own.
    pol = "Software\Policies\Microsoft\Office\" & Application.Version & "\Access\Security\Trusted Locations"
    common = "Software\Policies\Microsoft\Office\" & Application.Version & "\Common\Security\Trusted Locations"
    usr = "Software\Microsoft\Office\" & Application.Version & "\Access\Security\Trusted Locations"
    If RegDword(pol, "", "AllLocationsDisabled", 0) = 1 Then Exit Function
    If RegDword(common, "", "AllLocationsDisabled", 0) = 1 Then Exit Function
    If RegDword(usr, "", "AllLocationsDisabled", 0) = 1 Then Exit Function
    ' A policy setting wins over the user's own.
    allowNet = (RegDword(pol, "", "AllowNetworkLocations", RegDword(usr, "", "AllowNetworkLocations", 0)) = 1)
    isNet = (Left$(folder, 2) = "\\")
    If Not isNet And Mid$(folder, 2, 1) = ":" Then isNet = (GetDriveTypeW(StrPtr(Left$(folder, 3))) = DRIVE_REMOTE)
    If isNet And Not allowNet Then Exit Function
    onList = ListHas(pol, folder) Or ListHas(common, folder)
    ' "Allow mix of policy and user locations" turned off: Access ignores the
    ' user's own locations.
    If Not onList And RegDword(common, "", "Allow User Locations", 1) <> 0 Then onList = ListHas(usr, folder)
    InTrustedLocation = onList
    Exit Function
Fail:
    LogError "modTrust.InTrustedLocation", "folder " & folder
End Function

' True when one of the Location keys under root covers folder.
Private Function ListHas(ByVal root As String, ByVal folder As String) As Boolean
    Dim hKey As LongPtr, i As Long, nm As String, n As Long, p As String, subs As Boolean
    If RegOpenKeyExW(HKEY_CURRENT_USER, StrPtr(root), 0, KEY_READ, hKey) <> ERROR_SUCCESS Then Exit Function
    Do
        nm = String$(256, vbNullChar): n = 256
        If RegEnumKeyExW(hKey, i, StrPtr(nm), n, 0, 0, 0, 0) <> ERROR_SUCCESS Then Exit Do
        nm = Left$(nm, n)
        p = RegString(root, nm, "Path")
        If Len(p) > 0 Then
            subs = (RegDword(root, nm, "AllowSubfolders", 0) = 1)
            If PathCovers(p, subs, folder) Then ListHas = True: Exit Do
        End If
        i = i + 1
    Loop
    RegCloseKey hKey
End Function

' True when the Trusted Location loc (with subfolders when subs) covers
' folder. Both are compared normalized and ending in one backslash: a drive
' root (c:\) already has one. Public so the kit's smoke test can check it.
Public Function PathCovers(ByVal loc As String, ByVal subs As Boolean, ByVal folder As String) As Boolean
    loc = NormPath(loc): folder = NormPath(folder)
    If Right$(loc, 1) <> "\" Then loc = loc & "\"
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    PathCovers = (folder = loc) Or (subs And Left$(folder, Len(loc)) = loc)
End Function

' Lower case, environment variables expanded, no trailing backslash.
Private Function NormPath(ByVal p As String) As String
    p = LCase$(Trim$(ExpandEnv(p)))
    Do While Len(p) > 3 And Right$(p, 1) = "\"
        p = Left$(p, Len(p) - 1)
    Loop
    NormPath = p
End Function

' HKCU\root\subKey value name as text ("" when missing).
Private Function RegString(ByVal root As String, ByVal subKey As String, ByVal name As String) As String
    Dim hKey As LongPtr, kind As Long, cb As Long, buf As String
    If RegOpenKeyExW(HKEY_CURRENT_USER, StrPtr(KeyPath(root, subKey)), 0, KEY_READ, hKey) <> ERROR_SUCCESS Then Exit Function
    If RegQueryValueExW(hKey, StrPtr(name), 0, kind, 0, cb) = ERROR_SUCCESS And cb > 0 And kind <> REG_DWORD Then
        buf = String$(cb \ 2 + 1, vbNullChar)
        cb = Len(buf) * 2
        If RegQueryValueExW(hKey, StrPtr(name), 0, kind, StrPtr(buf), cb) = ERROR_SUCCESS Then
            CutAtNull buf
            RegString = buf
        End If
    End If
    RegCloseKey hKey
End Function

' Cut s at its first null character.
Private Sub CutAtNull(s As String)
    Dim i As Long
    i = InStr(s, vbNullChar)
    If i > 0 Then s = Left$(s, i - 1)
End Sub

' HKCU\root\subKey value name as a number (dflt when missing or not a DWORD).
Private Function RegDword(ByVal root As String, ByVal subKey As String, ByVal name As String, ByVal dflt As Long) As Long
    Dim hKey As LongPtr, kind As Long, cb As Long, v As Long
    RegDword = dflt
    If RegOpenKeyExW(HKEY_CURRENT_USER, StrPtr(KeyPath(root, subKey)), 0, KEY_READ, hKey) <> ERROR_SUCCESS Then Exit Function
    cb = 4
    If RegQueryValueExW(hKey, StrPtr(name), 0, kind, VarPtr(v), cb) = ERROR_SUCCESS Then
        If kind = REG_DWORD Then RegDword = v
    End If
    RegCloseKey hKey
End Function

Private Function KeyPath(ByVal root As String, ByVal subKey As String) As String
    KeyPath = root & IIf(Len(subKey) > 0, "\" & subKey, "")
End Function
