Option Compare Database
Option Explicit

' Access UI kit - "a newer version is available" notice.
'
' A release folder holds the app's current copy and a one-line text file,
' version.txt, naming that copy's version ("1.2.0"). At startup the app calls
' CheckForUpdate with its own version and the folder. When the file names a
' newer version, the user gets one KitMsgBox saying which version and where to
' get it, once per session (reopening the main window doesn't repeat it). A
' blank folder, a missing folder or file, or a file that can't be read never
' stops the app opening: the check is skipped, and a failed read is logged
' (LogError) without a message. An unreachable network folder can delay
' opening until Windows gives up on it (about 20 seconds).
'
' The app provides APP_NAME (named in the notice) and LogError (modErrors).

Public Const KIT_VERSION_FILE As String = "version.txt"

Private m_checked As Boolean      ' CheckForUpdate already ran this session

' Show the notice when the release folder's version.txt names a version newer
' than current. True when the notice was shown. Only the first call in a
' session checks; later calls return False. advice replaces the default next
' step ("Close <app> and copy the new version from: <folder>"), for an app
' where copying the new file over the old one would lose data.
Public Function CheckForUpdate(ByVal current As String, ByVal folder As String, _
        Optional ByVal advice As String = "") As Boolean
    Dim latest As String
    On Error GoTo Fail
    If m_checked Then Exit Function
    m_checked = True
    latest = ReleasedVersion(folder)
    If Len(latest) = 0 Then Exit Function
    If CompareVersions(latest, current) <= 0 Then Exit Function
    If Len(advice) = 0 Then advice = "Close " & APP_NAME & " and copy the new version from:" & vbCrLf & FolderPath(folder)
    KitMsgBox "Version " & latest & " of " & APP_NAME & " is available. You are using version " & current & "." & _
        vbCrLf & vbCrLf & advice, vbInformation, "New version available"
    CheckForUpdate = True
    Exit Function
Fail:
    LogError "modKitUpdate.CheckForUpdate", "folder " & folder & ", this version " & current
End Function

' The version named by folder\version.txt (its first non-blank line, trimmed),
' or "" when the folder is blank or the file doesn't exist. Errors reading the
' file go to the caller.
Public Function ReleasedVersion(ByVal folder As String) As String
    Dim path As String, line As Variant
    If Len(Trim$(folder)) = 0 Then Exit Function
    path = FolderPath(folder) & "\" & KIT_VERSION_FILE
    If Not FileExists(path) Then Exit Function
    For Each line In Split(Replace(ReadTextFile(path), vbCr, ""), vbLf)
        If Len(Trim$(line)) > 0 Then
            ReleasedVersion = Trim$(line)
            Exit Function
        End If
    Next
End Function

' Compare dotted versions part by part as numbers: -1 when a is older than b,
' 0 when they match, 1 when a is newer. A leading "v" is ignored, missing parts
' count as 0 ("1.2" = "1.2.0"), and each part counts only its leading digits
' ("3-beta" = 3; "x" = 0).
Public Function CompareVersions(ByVal a As String, ByVal b As String) As Integer
    Dim pa() As String, pb() As String, i As Long, n As Long, va As Double, vb As Double
    pa = Split(StripV(a), ".")
    pb = Split(StripV(b), ".")
    n = UBound(pa)
    If UBound(pb) > n Then n = UBound(pb)
    For i = 0 To n
        va = VersionPart(pa, i)
        vb = VersionPart(pb, i)
        If va < vb Then CompareVersions = -1: Exit Function
        If va > vb Then CompareVersions = 1: Exit Function
    Next
End Function

' A version without spaces or its leading "v".
Private Function StripV(ByVal s As String) As String
    s = Trim$(s)
    If LCase$(Left$(s, 1)) = "v" Then s = Mid$(s, 2)
    StripV = s
End Function

' Part i of a split version as a number: its leading digits, or 0 when the
' part is missing or doesn't start with a digit.
Private Function VersionPart(parts() As String, ByVal i As Long) As Double
    Dim s As String, k As Long
    If i > UBound(parts) Then Exit Function
    s = parts(i)
    For k = 1 To Len(s)
        If Not (Mid$(s, k, 1) Like "#") Then Exit For
    Next
    s = Left$(s, k - 1)
    If Len(s) > 0 Then VersionPart = Val(s)
End Function

' A folder path without spaces or a trailing backslash.
Private Function FolderPath(ByVal folder As String) As String
    folder = Trim$(folder)
    Do While Right$(folder, 1) = "\" And Len(folder) > 1
        folder = Left$(folder, Len(folder) - 1)
    Loop
    FolderPath = folder
End Function
